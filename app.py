# -*- coding: utf-8 -*-
"""
社内全在庫（弥生・本社・物流）確認ダッシュボード
====================================================================
  1. ZAICO ダウンロードCSV(.csv) をアップロード
  2. 物流(潤井戸)在庫CSV(.csv) をアップロード
  3. 弥生在庫表Excel(.xlsx) をアップロード
  4. 「エクセル作成」ボタンを1回押す
  5. 拠点別在庫一覧・弥生ゼロ在庫あぶり出しリストの2シート入りExcelをダウンロード

アップロードしたファイルはサーバーに保存されず、その場のメモリ上だけで
処理されます（ブラウザを閉じる／再実行するとリセットされます）。
====================================================================
"""

import streamlit as st

from inventory_core import (
    LOCATION_COLUMNS,
    ZAICO_LOCATION_MAP,
    build_summary_table,
)

st.set_page_config(
    page_title="社内全在庫（弥生・本社・物流）確認ダッシュボード",
    page_icon="📦",
    layout="wide",
)


# ============================================================
# パスワード保護（任意）
#   Secrets に APP_PASSWORD を設定した場合のみ、合言葉入力を求めます。
# ============================================================
def check_password() -> bool:
    if "APP_PASSWORD" not in st.secrets:
        return True

    def password_entered():
        if st.session_state.get("password") == st.secrets.get("APP_PASSWORD"):
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    st.title("📦 社内全在庫（弥生・本社・物流）確認ダッシュボード")
    st.text_input(
        "合言葉を入力してください",
        type="password",
        on_change=password_entered,
        key="password",
    )
    if st.session_state.get("password_correct") is False:
        st.error("合言葉が正しくありません。")
    return False


def main():
    if not check_password():
        return

    st.title("📦 社内全在庫（弥生・本社・物流）確認ダッシュボード")
    st.caption(
        "ZAICOのCSV・物流(潤井戸)在庫のCSV・弥生在庫表のExcelをアップロードして「エクセル作成」を押すと、"
        "拠点（本社 / 3F / 4F / 4F(PKG不備) / 物流 / 弥生）ごとの在庫数と、"
        "弥生ゼロ在庫あぶり出しリストをまとめたExcelが作成できます。"
    )

    with st.expander("このツールがやっていること", expanded=False):
        loc_lines = "\n".join(
            f"- ZAICO CSVで保管場所が **「{k}」** の数量合計 → **{v}** 列"
            for k, v in ZAICO_LOCATION_MAP.items()
        )
        st.markdown(
            f"""
品番（ZAICO「型番」＝物流「商品コード」＝弥生品番を変換したもの）で突き合わせ、1品番1行に集計します。

{loc_lines}
- 物流(潤井戸)在庫CSVの **「在庫数（総ピース）」** の合計 → **物流** 列
- 弥生在庫表の **G列（在庫数）** の合計 → **弥生** 列
- 弥生品番は末尾の `(数字)` を除去し先頭に `md-` を付与して正式品番に変換します（例: `76181-1(15438)` → `md-76181-1`）。
  変換ルールが当てはまらない品番（旧商品・外注品など）は **「要確認」** 列を付けて出力します。
- 同じ品番が複数行ある場合は数量を合算します。
- **合計** 列 ＝ 本社 ＋ 3F ＋ 4F ＋ 4F(PKG不備) ＋ 物流（**弥生は合計に含まず、合計の右隣に別掲**します）
- 並び順は品番の昇順（要確認品番は末尾）。いずれかの拠点に在庫データがある品番はすべて掲載します（合計0も含む）。
- 出力Excelは「拠点別在庫」「弥生ゼロ在庫あぶり出しリスト」の2シート構成です。
  あぶり出しリストは、弥生在庫が0（または存在しない）にもかかわらず4F(ZAICO)在庫が1以上ある品番の一覧です（要確認品番は対象外）。
            """
        )

    if "result" not in st.session_state:
        st.session_state.result = None

    # ------------------------------------------------------------
    st.subheader("① ZAICOダウンロードCSV をアップロード")
    zaico_upload = st.file_uploader(
        "ZAICOからダウンロードしたCSV(.csv) を1つ選択してください",
        type=["csv"],
        accept_multiple_files=False,
        key="zaico",
    )

    st.subheader("② 物流(潤井戸)在庫CSV をアップロード")
    butsuryu_upload = st.file_uploader(
        "物流の在庫リストCSV(.csv) を1つ選択してください",
        type=["csv"],
        accept_multiple_files=False,
        key="butsuryu",
    )

    st.subheader("③ 弥生在庫表 Excel をアップロード")
    yayoi_upload = st.file_uploader(
        "弥生在庫表のExcel(.xlsx) を1つ選択してください",
        type=["xlsx"],
        accept_multiple_files=False,
        key="yayoi",
    )

    st.subheader("④ エクセル作成")
    ready = zaico_upload is not None and butsuryu_upload is not None and yayoi_upload is not None
    run_button = st.button(
        "🧮 エクセル作成（拠点別在庫＋あぶり出しリストを作成）",
        type="primary",
        disabled=not ready,
        use_container_width=True,
    )
    if not ready:
        st.caption("※ ①②③すべてをアップロードすると押せるようになります。")

    # ------------------------------------------------------------
    if run_button:
        with st.status("集計しています...", expanded=True) as status:
            status.write(f"ZAICO CSV: {zaico_upload.name}")
            status.write(f"物流CSV: {butsuryu_upload.name}")
            status.write(f"弥生在庫表: {yayoi_upload.name}")
            result = build_summary_table(
                zaico_upload.read(), butsuryu_upload.read(), yayoi_upload.read()
            )
            st.session_state.result = result
            if result.success:
                status.update(label="✅ 集計が完了しました。", state="complete")
            else:
                status.update(label="❌ 集計を中止しました。", state="error")

    result = st.session_state.get("result")
    if result is None:
        st.info("①②③をアップロードして「エクセル作成」を押すと、ここに結果とダウンロードボタンが表示されます。")
        return

    if not result.success:
        st.error("❌ " + (result.error or "不明なエラーが発生しました。"))
        return

    for w in result.warnings:
        st.warning("⚠️ " + w)
    for i in result.infos:
        st.info("ℹ️ " + i)

    # --- サマリー ---
    st.subheader("結果サマリー")
    cols = st.columns(len(LOCATION_COLUMNS) + 3)
    cols[0].metric("品番数", f"{result.sku_count:,}")
    for i, loc in enumerate(LOCATION_COLUMNS, start=1):
        cols[i].metric(f"{loc} 在庫", f"{result.location_totals.get(loc, 0):,}")
    cols[-2].metric("要確認品番数", f"{result.needs_check_count:,}")
    cols[-1].metric("あぶり出し件数", f"{result.abrid_count:,}")

    # --- ダウンロード ---
    st.subheader("⑤ ダウンロード")
    st.download_button(
        f"📥 {result.output_name}",
        data=result.output_bytes,
        file_name=result.output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    # --- プレビュー ---
    st.subheader("内容の確認（画面プレビュー）")
    st.caption("※ ダウンロードするExcelと同じ内容です（シート①最下行に合計行が付きます）。")

    tab1, tab2 = st.tabs(["① 拠点別在庫一覧", "② 弥生ゼロ在庫あぶり出しリスト"])

    with tab1:
        df = result.table
        q1 = st.text_input("品番・弥生品番・商品名で絞り込み（部分一致）", "", key="q1")
        view = df
        if q1.strip():
            key = q1.strip()
            view = df[
                df["品番"].str.contains(key, case=False, na=False)
                | df["弥生品番"].str.contains(key, case=False, na=False)
                | df["商品名"].str.contains(key, case=False, na=False)
            ]
        st.dataframe(
            view,
            use_container_width=True,
            hide_index=True,
            column_config={
                c: st.column_config.NumberColumn(format="%d")
                for c in LOCATION_COLUMNS + ["合計"]
            },
        )
        st.caption(f"表示 {len(view):,} / {len(df):,} 品番（うち要確認 {result.needs_check_count:,} 品番）")

    with tab2:
        df2 = result.abrid_table
        q2 = st.text_input("品番・弥生品番・商品名で絞り込み（部分一致）", "", key="q2")
        view2 = df2
        if q2.strip():
            key = q2.strip()
            view2 = df2[
                df2["品番"].str.contains(key, case=False, na=False)
                | df2["弥生品番"].str.contains(key, case=False, na=False)
                | df2["商品名"].str.contains(key, case=False, na=False)
            ]
        st.dataframe(
            view2,
            use_container_width=True,
            hide_index=True,
            column_config={
                c: st.column_config.NumberColumn(format="%d")
                for c in ["弥生在庫数", "4F在庫数", "4F(PKG不備)在庫数"]
            },
        )
        st.caption(f"表示 {len(view2):,} / {len(df2):,} 品番")


if __name__ == "__main__":
    main()
