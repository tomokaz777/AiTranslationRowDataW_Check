import streamlit as st
import pandas as pd
from checker import TranslationChecker
from excel_handler import read_excel, write_excel
import os
from dotenv import load_dotenv
import time
import threading

load_dotenv()

st.set_page_config(
    page_title="翻訳品質チェックツール",
    page_icon="📋",
    layout="wide",
)

# セッション状態の初期化
for key, default in [
    ("result_df", None),
    ("original_filename", ""),
    ("is_checked", False),
    ("output_bytes", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# アニメーション用スピナーフレーム
SPIN_FRAMES = ["◐", "◓", "◑", "◒"]


def fmt_time(sec: float) -> str:
    """秒数を M:SS または H:MM:SS 形式に変換"""
    sec = max(0, int(sec))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


# ─────────────────────────────────────────
# サイドバー
# ─────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ 設定")
    st.markdown("製品マニュアルの AI 翻訳品質を Claude で自動チェックします。")
    st.divider()

    # APIキー
    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        st.success("✅ 環境変数から API キーを取得済み")
        api_key = env_key
    else:
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-...",
            help="https://console.anthropic.com でキーを取得してください",
        )
        if not api_key:
            st.error("⚠️ API キーを入力してください")

    st.divider()

    # 並行処理数（上限30に拡大）
    concurrency = st.slider(
        "並行処理数",
        min_value=1,
        max_value=30,
        value=3,
        help=(
            "同時に Claude API へリクエストする数。\n"
            "多いほど速いがレート制限エラーが発生しやすくなります。\n"
            "推奨: 3（安定）, 5（Tier1）, 15（Tier2以上）"
        ),
    )

    st.divider()

    # PASSスキップ設定
    skip_pass = st.checkbox(
        "PASS行をスキップ（高速化）",
        value=True,
        help=(
            "ONにすると、既存評価がPASSの行はAPIを呼ばずにそのままPASSとして扱います。\n"
            "大量のPASS行がある場合に大幅に高速化されます。\n"
            "厳密な再チェックが必要な場合はOFFにしてください。"
        ),
    )

    if skip_pass:
        st.info("PASSスキップ: ON\nFAIL行のみAPIを呼び出します")
    else:
        st.warning("PASSスキップ: OFF\n全行APIを呼び出します（低速）")

    st.divider()

    with st.expander("📖 使い方ガイド"):
        st.markdown(
            """
1. **API キー**を入力（または環境変数 `ANTHROPIC_API_KEY` を設定）
2. **Excel ファイル**をアップロード
3. **「品質チェック開始」**をクリック
4. 完了後、**結果をダウンロード**

**速度の目安（PASSスキップON時）:**
- FAIL率13%・並行数10: ~2分 / 1,000行
- FAIL率13%・並行数20: ~8〜9分 / 20,000行

**入力列の構成:**
- D列: 日本語原文
- E列: AI 英訳文
- L列: 評価① (PASS/FAIL)
- M-P列: 評価詳細（FAIL時参照）

**出力列:**
- Q列: Arizo品質チェック結果 (PASS/FAIL)
- R列: Arizo推奨英訳（FAILの場合のみ）

---

**【仕様】Pass / Fail の処理フロー:**
- **L列が PASS** : D列（日本語原文）と E列（AI英訳）を Claude が比較・再確認（PASSスキップONの場合は API スキップ）→ Q列に PASS/FAIL、R列は空
- **L列が FAIL** : D列・E列に加え N〜P列（既存の指摘内容）も参照して Claude が評価・修正提案 → Q列に PASS/FAIL、R列に推奨英訳（FAILの場合のみ）
- **日本語列が空白** : 処理スキップ → Q列・R列ともに空
            """
        )


# ─────────────────────────────────────────
# メインエリア
# ─────────────────────────────────────────
st.title("📋 翻訳品質チェックツール")

# ── 1. ファイルアップロード ──────────────
st.header("1. ファイルアップロード")
uploaded_file = st.file_uploader(
    "TTより納品された生データExcel ファイルをアップロード（.xlsx / .xls）。詳細は使い方ガイド参照。",
    type=["xlsx", "xls"],
    help="ドラッグ＆ドロップまたはクリックでファイルを選択",
)

if uploaded_file is not None:
    st.session_state.original_filename = uploaded_file.name

    try:
        df = read_excel(uploaded_file)
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました: {e}")
        st.stop()

    # ── 2. データプレビュー ──────────────
    st.header("2. データプレビュー")

    cols = df.columns.tolist()
    d_col = cols[3] if len(cols) > 3 else None
    l_col = cols[11] if len(cols) > 11 else None

    total_rows = len(df)
    check_target = 0
    pass_count_pre = 0
    fail_count_pre = 0

    if d_col:
        check_target = int(
            (df[d_col].astype(str).str.strip().replace("nan", "") != "").sum()
        )
    if l_col:
        pass_count_pre = int(
            (df[l_col].astype(str).str.upper().str.strip() == "PASS").sum()
        )
        fail_count_pre = int(
            (df[l_col].astype(str).str.upper().str.strip() == "FAIL").sum()
        )

    # 推定処理時間を表示
    if skip_pass:
        api_calls = fail_count_pre
        est_sec = (api_calls / max(concurrency, 1)) * 4
    else:
        api_calls = check_target
        est_sec = (api_calls / max(concurrency, 1)) * 2.5

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("総行数", f"{total_rows:,}")
    with c2:
        st.metric("チェック対象行数", f"{check_target:,}")
    with c3:
        st.metric("PASS 数（既存）", f"{pass_count_pre:,}")
    with c4:
        st.metric("FAIL 数（既存）", f"{fail_count_pre:,}")
    with c5:
        st.metric("API呼び出し予定", f"{api_calls:,} 行", help="PASSスキップ設定により変動")

    st.caption(f"⏱ 推定処理時間: 約 **{fmt_time(est_sec)}**（並行数{concurrency}・平均4秒/回）")

    preview_cols = cols[:12] if len(cols) >= 12 else cols
    st.dataframe(df[preview_cols].head(20), use_container_width=True)

    # ── 3. 実行ボタン ──────────────────────
    st.header("3. 品質チェック実行")

    if not api_key:
        st.warning("API キーが設定されていません。サイドバーで入力してください。")
        st.stop()

    run_button = st.button(
        "▶ 品質チェック開始",
        type="primary",
        use_container_width=True,
        disabled=st.session_state.is_checked,
    )

    if run_button:
        st.session_state.is_checked = False
        st.session_state.result_df = None
        st.session_state.output_bytes = None

        # 行データをリストに変換
        rows_data = [
            {
                "japanese":      str(row.iloc[3])  if len(row) > 3  else "",
                "ai_translation": str(row.iloc[4]) if len(row) > 4  else "",
                "eval_l":        str(row.iloc[11]) if len(row) > 11 else "",
                "eval_m":        str(row.iloc[12]) if len(row) > 12 else "",
                "eval_n":        str(row.iloc[13]) if len(row) > 13 else "",
                "eval_o":        str(row.iloc[14]) if len(row) > 14 else "",
                "eval_p":        str(row.iloc[15]) if len(row) > 15 else "",
            }
            for _, row in df.iterrows()
        ]

        # ── 4. 進捗表示 ──────────────────────
        st.header("4. 処理中...")

        # プレースホルダーを事前に確保
        ph_header   = st.empty()   # スピナー + タイトル行
        ph_bar      = st.empty()   # プログレスバー
        ph_metrics  = st.empty()   # 速度 / ETA / 経過時間
        ph_counts   = st.empty()   # PASS / FAIL カウンター

        # 共有状態
        progress_state = {
            "done": 0, "total": len(rows_data),
            "pass": 0, "fail": 0,
        }

        def progress_callback(done, total, pass_cnt, fail_cnt):
            progress_state["done"] = done
            progress_state["pass"] = pass_cnt
            progress_state["fail"] = fail_cnt

        results_container: list = []
        error_container:   list = []

        def run_check():
            try:
                checker = TranslationChecker(
                    api_key=api_key,
                    concurrency=concurrency,
                    skip_pass=skip_pass,
                )
                results = checker.check_batch(rows_data, progress_callback=progress_callback)
                results_container.extend(results)
            except Exception as e:
                error_container.append(str(e))

        thread = threading.Thread(target=run_check, daemon=True)
        thread.start()
        start_time = time.time()
        spin_idx = 0

        # ── ポーリングループ（0.2秒更新）──
        while thread.is_alive():
            done  = progress_state["done"]
            total = progress_state["total"]
            p     = progress_state["pass"]
            f     = progress_state["fail"]

            elapsed    = time.time() - start_time
            pct        = done / total if total > 0 else 0
            rows_sec   = done / elapsed if elapsed > 0 and done > 0 else 0
            remaining  = total - done
            eta_sec    = remaining / rows_sec if rows_sec > 0 else 0

            spin_char  = SPIN_FRAMES[spin_idx % len(SPIN_FRAMES)]
            spin_idx  += 1

            # ヘッダー行: スピナー + 進捗率
            ph_header.markdown(
                f"### {spin_char} 処理中...　"
                f"**{pct * 100:.1f}%** 完了"
            )

            # プログレスバー
            ph_bar.progress(
                pct,
                text=f"{done:,} / {total:,} 行",
            )

            # 速度・ETA・経過時間
            speed_str = f"{rows_sec:.1f} 行/秒" if rows_sec > 0 else "計算中..."
            eta_str   = fmt_time(eta_sec)         if rows_sec > 0 else "計算中..."
            ph_metrics.info(
                f"⏱ 経過: **{fmt_time(elapsed)}**　｜　"
                f"処理速度: **{speed_str}**　｜　"
                f"残り推定: **{eta_str}**"
            )

            # PASS / FAIL カウンター
            error_cnt = done - p - f
            ph_counts.markdown(
                f"✅ PASS: **{p:,}**　　"
                f"❌ FAIL: **{f:,}**　　"
                f"⏳ 未完了: **{remaining:,}**"
                + (f"　　⚠️ ERROR: **{error_cnt:,}**" if error_cnt > 0 else "")
            )

            time.sleep(0.2)

        thread.join()

        # 完了時に100%表示
        ph_bar.progress(1.0, text=f"{len(rows_data):,} / {len(rows_data):,} 行")
        ph_header.markdown("### ✅ 処理完了！")

        if error_container:
            st.error(f"エラーが発生しました: {error_container[0]}")
            st.stop()

        results = results_container
        elapsed_total = time.time() - start_time

        # 完了サマリー
        final_pass = sum(1 for r in results if r.get("result") == "PASS")
        final_fail = sum(1 for r in results if r.get("result") == "FAIL")
        ph_metrics.success(
            f"完了！　処理時間: **{fmt_time(elapsed_total)}**　｜　"
            f"PASS: **{final_pass:,}**　FAIL: **{final_fail:,}**　"
            f"（{len(rows_data):,} 行）"
        )
        ph_counts.empty()

        # 結果をDataFrameに追加
        result_df = df.copy()
        result_df["Arizo品質チェック結果"] = [r.get("result", "") for r in results]
        result_df["Arizo推奨英訳"]        = [
            r.get("suggested", "") if r.get("result") != "PASS" else ""
            for r in results
        ]

        st.session_state.result_df        = result_df
        st.session_state.is_checked       = True
        st.session_state.output_bytes     = write_excel(
            result_df, st.session_state.original_filename
        )

        st.rerun()

    # ── 5. 結果テーブル ────────────────────
    if st.session_state.is_checked and st.session_state.result_df is not None:
        st.header("5. チェック結果")

        result_df = st.session_state.result_df
        q_col = "Arizo品質チェック結果"

        def highlight_row(row):
            val = str(row.get(q_col, "")).upper() if q_col in row.index else ""
            if val == "PASS":
                return ["background-color: #E2EFDA"] * len(row)
            elif val == "FAIL":
                return ["background-color: #FCE4D6"] * len(row)
            return [""] * len(row)

        st.dataframe(
            result_df.style.apply(highlight_row, axis=1),
            use_container_width=True,
        )

        # ── 6. ダウンロードボタン ────────────
        st.header("6. ダウンロード")
        base_name = os.path.splitext(st.session_state.original_filename)[0]
        st.download_button(
            label="⬇️ 結果をダウンロード（Excel）",
            data=st.session_state.output_bytes,
            file_name=f"{base_name}_QC済み.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.info("👆 Excel ファイルをアップロードしてください（.xlsx / .xls）")

    if st.session_state.is_checked:
        st.session_state.is_checked   = False
        st.session_state.result_df    = None
        st.session_state.output_bytes = None
