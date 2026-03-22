import anthropic
import asyncio
import json
import re


class TranslationChecker:
    def __init__(self, api_key: str, concurrency: int = 10, skip_pass: bool = True):
        self.api_key = api_key
        self.concurrency = concurrency
        self.skip_pass = skip_pass
        self.model = "claude-sonnet-4-20250514"
        self.max_tokens = 1000

    # ──────────────────────────────────────────
    # 公開API
    # ──────────────────────────────────────────

    def check_batch(self, rows: list[dict], progress_callback=None) -> list[dict]:
        """
        複数行を非同期並行処理でチェック（スレッドから呼び出し可）
        progress_callback(done, total, pass_count, fail_count) を都度呼び出す
        """
        return asyncio.run(self._check_batch_async(rows, progress_callback))

    # ──────────────────────────────────────────
    # 非同期処理コア
    # ──────────────────────────────────────────

    async def _check_batch_async(
        self, rows: list[dict], progress_callback=None
    ) -> list[dict]:
        client = anthropic.AsyncAnthropic(api_key=self.api_key)
        semaphore = asyncio.Semaphore(self.concurrency)
        results = [None] * len(rows)
        counters = {"done": 0, "pass": 0, "fail": 0}
        lock = asyncio.Lock()

        async def process(idx: int, row: dict):
            result = await self._check_row_async(client, row, semaphore)
            async with lock:
                results[idx] = result
                counters["done"] += 1
                r = result.get("result", "")
                if r == "PASS":
                    counters["pass"] += 1
                elif r == "FAIL":
                    counters["fail"] += 1
                if progress_callback:
                    progress_callback(
                        counters["done"], len(rows),
                        counters["pass"], counters["fail"],
                    )

        await asyncio.gather(*[process(i, row) for i, row in enumerate(rows)])
        return results

    async def _check_row_async(
        self,
        client: anthropic.AsyncAnthropic,
        row: dict,
        semaphore: asyncio.Semaphore,
    ) -> dict:
        japanese = str(row.get("japanese", "")).strip()
        ai_translation = str(row.get("ai_translation", "")).strip()
        eval_l = str(row.get("eval_l", "")).strip().upper()

        if not japanese or japanese.lower() == "nan":
            return {"result": "", "suggested": ""}

        # PASSスキップモード: API呼び出しなしで即座にPASS返却
        if self.skip_pass and eval_l == "PASS":
            return {"result": "PASS", "suggested": ""}

        for attempt in range(2):
            try:
                async with semaphore:
                    if eval_l == "PASS":
                        result_str = await self._call_pass_async(
                            client, japanese, ai_translation
                        )
                        return {"result": result_str, "suggested": ""}
                    else:
                        return await self._call_fail_async(
                            client,
                            japanese,
                            ai_translation,
                            str(row.get("eval_m", "")).strip(),
                            str(row.get("eval_n", "")).strip(),
                            str(row.get("eval_o", "")).strip(),
                            str(row.get("eval_p", "")).strip()[:500],
                        )
            except anthropic.RateLimitError:
                # レート制限: 少し長めに待つ
                wait = 10 if attempt == 0 else 30
                await asyncio.sleep(wait)
            except Exception as e:
                if attempt == 0:
                    await asyncio.sleep(3)
                else:
                    return {"result": "ERROR", "suggested": str(e)}

        return {"result": "ERROR", "suggested": "Max retries exceeded"}

    # ──────────────────────────────────────────
    # API呼び出し
    # ──────────────────────────────────────────

    async def _call_pass_async(
        self, client: anthropic.AsyncAnthropic, japanese: str, ai_translation: str
    ) -> str:
        response = await client.messages.create(
            model=self.model,
            max_tokens=10,
            temperature=0,
            system="あなたはプロの日本語→英語翻訳品質チェッカーです。製品マニュアルの翻訳を評価してください。",
            messages=[
                {"role": "user", "content": self._build_prompt_pass(japanese, ai_translation)}
            ],
        )
        text = response.content[0].text.strip().upper()
        return "FAIL" if "FAIL" in text else "PASS"

    async def _call_fail_async(
        self,
        client: anthropic.AsyncAnthropic,
        japanese: str,
        ai_translation: str,
        eval_m: str,
        eval_n: str,
        eval_o: str,
        eval_p: str,
    ) -> dict:
        response = await client.messages.create(
            model=self.model,
            max_tokens=self.max_tokens,
            temperature=0,
            system="あなたはプロの日本語→英語翻訳品質チェッカーです。製品マニュアルの翻訳を評価・修正してください。",
            messages=[
                {
                    "role": "user",
                    "content": self._build_prompt_fail(
                        japanese, ai_translation, eval_m, eval_n, eval_o, eval_p
                    ),
                }
            ],
        )
        return self._parse_fail_response(response.content[0].text.strip())

    # ──────────────────────────────────────────
    # プロンプト構築
    # ──────────────────────────────────────────

    def _build_prompt_pass(self, japanese: str, ai_translation: str) -> str:
        return f"""日本語原文: "{japanese}"
英訳: "{ai_translation}"

この英訳が日本語原文を正確に翻訳しているか判定してください。
判定基準:
- 製品コード・型番・記号など翻訳不要の項目はPASS
- 意味・情報が正確に伝わっていればPASS
- 意味の欠落・誤訳・不自然な表現があればFAIL

"PASS" または "FAIL" の1単語のみで回答してください。"""

    def _build_prompt_fail(
        self,
        japanese: str,
        ai_translation: str,
        eval_m: str,
        eval_n: str,
        eval_o: str,
        eval_p: str,
    ) -> str:
        return f"""日本語原文: "{japanese}"
現在の英訳（問題あり）: "{ai_translation}"

この翻訳には以下の問題が指摘されています:
- 評価グレード: {eval_m}
- 指摘内容: {eval_n}
- 詳細: {eval_o}
- 補足: {eval_p}

タスク:
1. 修正後の英訳が正確かどうか最終判定（PASS/FAIL）
2. 上記の問題を修正した、正確な英訳を提案する

必ず以下のJSON形式のみで回答すること（マークダウン不要）:
{{"result": "PASS", "translation": "corrected English translation here"}}"""

    # ──────────────────────────────────────────
    # レスポンスパース
    # ──────────────────────────────────────────

    def _parse_fail_response(self, text: str) -> dict:
        """JSONパース。失敗時は正規表現でフォールバック"""
        for src in [text, re.sub(r"```(?:json)?\s*|\s*```", "", text).strip()]:
            try:
                data = json.loads(src)
                result = str(data.get("result", "FAIL")).upper()
                if result not in ("PASS", "FAIL"):
                    result = "FAIL"
                # PASSの場合は推奨英訳不要
                suggested = "" if result == "PASS" else data.get("translation", "")
                return {"result": result, "suggested": suggested}
            except json.JSONDecodeError:
                pass

        result_match = re.search(r'"result"\s*:\s*"(PASS|FAIL)"', text, re.IGNORECASE)
        translation_match = re.search(
            r'"translation"\s*:\s*"((?:[^"\\]|\\.)*)"', text, re.DOTALL
        )
        result = result_match.group(1).upper() if result_match else "FAIL"
        return {
            "result": result,
            # PASSの場合は推奨英訳不要
            "suggested": (
                translation_match.group(1).replace('\\"', '"')
                if translation_match and result == "FAIL"
                else ""
            ),
        }
