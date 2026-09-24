# -*- coding: utf-8 -*-
"""
社内全在庫（弥生・本社・物流）集計ロジック
====================================================================
やっていること:

  ・ZAICO ダウンロードCSV の「型番」「保管場所」「数量」「物品名」
  ・物流(潤井戸)在庫CSV の「商品コード」「商品名」「在庫数（総ピース）」
  ・弥生在庫表Excel の A列(品番) / B列(商品名) / G列(在庫数)

を品番（弥生品番は末尾の括弧を除去し「md-」を付与して正式品番へ変換）で
突き合わせ、拠点ごとの在庫数を1行にまとめたExcel（2シート）を作成します。

  出力列    集計元
  --------  ------------------------------------------------
  本社      ZAICO 保管場所「本社在庫」の数量合計
  3F        ZAICO 保管場所「★3F EC★」の数量合計
  4F        ZAICO 保管場所「◆4F EC◆」「□４F 他社商品□」の数量合計
  4F(PKG不備) ZAICO 保管場所「●４F PKG不備●」の数量合計
  物流      物流CSV「在庫数（総ピース）」の合計
  弥生      弥生在庫表 G列（在庫数）の合計

シート②「弥生ゼロ在庫あぶり出しリスト」は、弥生在庫が0（または弥生に存在しない）
にもかかわらず4F(ZAICO)在庫が1以上ある品番を抽出したものです。

アップロードされたファイルのバイト列を受け取り、結果もバイト列で返します
（サーバーには保存されません）。
====================================================================
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ====================================================================
# 業務ルール設定
#   運用ルールが変わったときは、基本ここだけ直せば済むようにまとめています。
# ====================================================================

CSV_ENCODINGS_TO_TRY = ("utf-8-sig", "utf-8", "cp932", "shift_jis")
FULLWIDTH_SPACE = "　"

# --- ZAICO CSV ---
ZAICO_COL_SKU = "型番"
ZAICO_COL_LOCATION = "保管場所"
ZAICO_COL_QTY = "数量"
ZAICO_COL_NAME = "物品名"

# ZAICOの「保管場所」→ 拠点列（この4つ以外の保管場所は集計対象外＝警告表示）
ZAICO_LOCATION_MAP = {
    "本社在庫": "本社",
    "★3F EC★": "3F",
    "◆4F EC◆": "4F",
    "□４F 他社商品□": "4F",       # 他社商品在庫は4F列へ合算
    "●４F PKG不備●": "4F(PKG不備)",
}

# --- 物流(潤井戸)在庫 CSV ---
BUTSURYU_COL_CODE = "商品コード"
BUTSURYU_COL_NAME = "商品名"
BUTSURYU_COL_QTY = "在庫数（総ピース）"

# --- 弥生在庫表 Excel（ヘッダー行なし。1シート目のA/B/G列を位置参照） ---
YAYOI_COL_IDX_SKU = 0    # A列: 弥生品番
YAYOI_COL_IDX_NAME = 1   # B列: 商品名
YAYOI_COL_IDX_QTY = 6    # G列: 在庫数

# 弥生品番の変換ルール: 末尾の "(数字)" を除去し、先頭に "md-" を付与
#   例: "76181-1(15438)" -> "md-76181-1"
YAYOI_SKU_PATTERN = re.compile(r"^(.+)\((\d+)\)$")
YAYOI_SKU_PREFIX = "md-"

# --- 出力 ---
# 合計に含める拠点列（E~I列）。弥生はここに含めず、合計の後ろ（K列）に別掲する。
PHYSICAL_LOCATION_COLUMNS = ["本社", "3F", "4F", "4F(PKG不備)", "物流"]
LOCATION_COLUMNS = PHYSICAL_LOCATION_COLUMNS + ["弥生"]   # 表示順（画面サマリー等で使用）
SHEET1_COLUMNS = (
    ["品番", "弥生品番", "要確認", "商品名"] + PHYSICAL_LOCATION_COLUMNS + ["合計", "弥生"]
)
SHEET2_COLUMNS = ["品番", "弥生品番", "商品名", "弥生在庫数", "4F在庫数", "4F(PKG不備)在庫数"]

SHEET1_NAME = "拠点別在庫"
SHEET2_NAME = "弥生ゼロ在庫あぶり出しリスト"
OUTPUT_PREFIX = "社内全在庫_弥生本社物流"

NO_SKU_MARK = "―"
NEEDS_CHECK_MARK = "要確認"


# ====================================================================
# データ構造
# ====================================================================

@dataclass
class BuildResult:
    success: bool = False
    error: Optional[str] = None
    output_name: Optional[str] = None
    output_bytes: Optional[bytes] = None

    table: Optional[pd.DataFrame] = None            # シート①プレビュー用（合計行を含まない）
    abrid_table: Optional[pd.DataFrame] = None       # シート②プレビュー用

    location_totals: dict = field(default_factory=dict)   # {拠点: 合計数量}
    grand_total: int = 0
    sku_count: int = 0
    needs_check_count: int = 0
    abrid_count: int = 0

    # 注意喚起
    skipped_locations: dict = field(default_factory=dict)     # 集計対象外の保管場所 {名前: 行数}
    zaico_missing_sku_rows: int = 0                            # 型番が空のZAICO行数
    zaico_missing_sku_qty: int = 0                              # その数量合計
    warnings: list = field(default_factory=list)
    infos: list = field(default_factory=list)


# ====================================================================
# 共通ユーティリティ
# ====================================================================

def _normalize_key(value) -> str:
    """品番（型番／商品コード）を突き合わせできるよう正規化する。"""
    if value is None:
        return ""
    s = str(value).strip().replace(FULLWIDTH_SPACE, "").strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s


def _read_csv_any_encoding(csv_bytes: bytes) -> pd.DataFrame:
    last_err: Optional[Exception] = None
    for enc in CSV_ENCODINGS_TO_TRY:
        try:
            df = pd.read_csv(io.BytesIO(csv_bytes), encoding=enc, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
        except Exception as e:  # pd.errors.ParserError など
            last_err = e
            continue
    raise RuntimeError(f"CSVを読み込めませんでした（文字コード判定に失敗）: {last_err}")


def _require_columns(df: pd.DataFrame, cols, label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise RuntimeError(
            f"{label}に必要な列が見つかりません: {missing}\n"
            f"読み込んだ列: {list(df.columns)}"
        )


def convert_yayoi_sku(raw: str) -> Optional[str]:
    """弥生品番を正式品番に変換する。変換できなければ None（要確認扱い）。"""
    m = YAYOI_SKU_PATTERN.match(raw.strip())
    if not m:
        return None
    return YAYOI_SKU_PREFIX + m.group(1).strip()


# ====================================================================
# 集計
# ====================================================================

def _aggregate_zaico(df: pd.DataFrame, result: BuildResult) -> pd.DataFrame:
    """ZAICO CSV → 品番ごとの {本社, 3F, 4F, 4F(PKG不備)} 在庫数（＋ 商品名候補）"""
    _require_columns(
        df, [ZAICO_COL_SKU, ZAICO_COL_LOCATION, ZAICO_COL_QTY, ZAICO_COL_NAME],
        "ZAICOのCSV",
    )

    work = pd.DataFrame({
        "sku": df[ZAICO_COL_SKU].map(_normalize_key),
        "loc_raw": df[ZAICO_COL_LOCATION].astype(str).str.strip(),
        "qty": pd.to_numeric(df[ZAICO_COL_QTY], errors="coerce").fillna(0),
        "name": df[ZAICO_COL_NAME].astype(str).str.strip(),
    })
    work["拠点"] = work["loc_raw"].map(ZAICO_LOCATION_MAP)

    # 型番が空の行（＝品番で突き合わせできない行）
    blank = work[work["sku"] == ""]
    result.zaico_missing_sku_rows = int(len(blank))
    result.zaico_missing_sku_qty = int(blank["qty"].sum())

    # 集計対象外の保管場所
    skipped = work[work["拠点"].isna() & (work["loc_raw"] != "")]
    if not skipped.empty:
        result.skipped_locations = (
            skipped["loc_raw"].value_counts().astype(int).to_dict()
        )

    valid = work[(work["sku"] != "") & work["拠点"].notna()].copy()

    # 拠点別ピボット（同じ拠点列に複数の保管場所が合算されるケースがあるため sum で合算）
    pivot = (
        valid.pivot_table(index="sku", columns="拠点", values="qty",
                          aggfunc="sum", fill_value=0)
        .reindex(columns=["本社", "3F", "4F", "4F(PKG不備)"], fill_value=0)
    )
    pivot = pivot.astype("int64")

    # 品番ごとの商品名（最頻値。無ければ空）
    name_map = (
        valid[valid["name"] != ""]
        .groupby("sku")["name"]
        .agg(lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0])
    )
    pivot["_zaico_name"] = name_map
    pivot.index.name = "品番"
    return pivot.reset_index()


def _aggregate_butsuryu(df: pd.DataFrame, result: BuildResult) -> pd.DataFrame:
    """物流CSV → 品番ごとの 物流在庫数（＋ 商品名候補）"""
    _require_columns(
        df, [BUTSURYU_COL_CODE, BUTSURYU_COL_NAME, BUTSURYU_COL_QTY],
        "物流(潤井戸)在庫のCSV",
    )

    work = pd.DataFrame({
        "sku": df[BUTSURYU_COL_CODE].map(_normalize_key),
        "qty": pd.to_numeric(df[BUTSURYU_COL_QTY], errors="coerce").fillna(0),
        "name": df[BUTSURYU_COL_NAME].astype(str).str.strip(),
    })
    work = work[work["sku"] != ""]

    grouped = work.groupby("sku", as_index=False).agg(
        物流=("qty", "sum"),
        _butsuryu_name=("name", lambda s: s.iloc[0] if len(s) else ""),
    )
    grouped["物流"] = grouped["物流"].astype("int64")
    grouped = grouped.rename(columns={"sku": "品番"})
    return grouped


def _read_yayoi_excel(xlsx_bytes: bytes) -> pd.DataFrame:
    try:
        df = pd.read_excel(
            io.BytesIO(xlsx_bytes), sheet_name=0, header=None, engine="openpyxl"
        )
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"弥生在庫表Excelの読み込みに失敗しました。\n{e}") from e

    if df.shape[1] <= YAYOI_COL_IDX_QTY:
        raise RuntimeError(
            f"弥生在庫表Excelの列数が想定より少なく、G列（在庫数）を読み取れません。"
            f"読み込んだ列数: {df.shape[1]}"
        )
    return df


def _aggregate_yayoi(xlsx_bytes: bytes, result: BuildResult):
    """弥生在庫表Excel → (変換済み品番ごとの弥生在庫集計, 要確認品番ごとの弥生在庫集計)"""
    raw_df = _read_yayoi_excel(xlsx_bytes)

    work = pd.DataFrame({
        "raw_sku": raw_df[YAYOI_COL_IDX_SKU].map(_normalize_key),
        "name": raw_df[YAYOI_COL_IDX_NAME].astype(str).str.strip().replace("nan", ""),
        "qty": pd.to_numeric(raw_df[YAYOI_COL_IDX_QTY], errors="coerce").fillna(0),
    })
    work = work[work["raw_sku"] != ""]

    work["品番"] = work["raw_sku"].map(convert_yayoi_sku)

    converted = work[work["品番"].notna()].copy()
    needs_check = work[work["品番"].isna()].copy()

    converted_agg = converted.groupby("品番", as_index=False).agg(
        弥生=("qty", "sum"),
        _yayoi_raw_sku=("raw_sku", "first"),
        _yayoi_name=("name", lambda s: next((v for v in s if v), "")),
    )
    converted_agg["弥生"] = converted_agg["弥生"].astype("int64")

    needs_check_agg = needs_check.groupby("raw_sku", as_index=False).agg(
        弥生=("qty", "sum"),
        商品名=("name", lambda s: next((v for v in s if v), "")),
    )
    needs_check_agg["弥生"] = needs_check_agg["弥生"].astype("int64")
    needs_check_agg = needs_check_agg.rename(columns={"raw_sku": "弥生品番"})

    result.needs_check_count = int(len(needs_check_agg))
    if result.needs_check_count:
        result.infos.append(
            f"弥生在庫表のうち {result.needs_check_count} 件の品番は「md-＋末尾括弧除去」の"
            "変換ルールが当てはまらず「要確認」として出力しました（旧商品・外注品など）。"
        )

    return converted_agg, needs_check_agg


def build_summary_table(zaico_bytes: bytes, butsuryu_bytes: bytes, yayoi_bytes: bytes) -> BuildResult:
    result = BuildResult()
    try:
        zaico_df = _read_csv_any_encoding(zaico_bytes)
    except Exception as e:  # noqa: BLE001
        result.error = f"ZAICO CSV の読み込みに失敗しました。\n{e}"
        return result
    try:
        butsuryu_df = _read_csv_any_encoding(butsuryu_bytes)
    except Exception as e:  # noqa: BLE001
        result.error = f"物流(潤井戸)在庫CSV の読み込みに失敗しました。\n{e}"
        return result

    try:
        zaico_agg = _aggregate_zaico(zaico_df, result)
        butsuryu_agg = _aggregate_butsuryu(butsuryu_df, result)
        yayoi_agg, yayoi_needs_check = _aggregate_yayoi(yayoi_bytes, result)
    except Exception as e:  # noqa: BLE001
        result.error = str(e)
        return result

    merged = pd.merge(zaico_agg, butsuryu_agg, on="品番", how="outer")
    merged = pd.merge(merged, yayoi_agg, on="品番", how="outer")

    for col in LOCATION_COLUMNS:
        if col not in merged.columns:
            merged[col] = 0
        merged[col] = merged[col].fillna(0).astype("int64")

    # 商品名: ZAICO 物品名 → 物流 商品名 → 弥生 商品名 の優先順
    def _clean_name_series(s):
        if s is None:
            return pd.Series([""] * len(merged))
        return s.fillna("").replace({"nan": ""})

    z = _clean_name_series(merged.get("_zaico_name"))
    b = _clean_name_series(merged.get("_butsuryu_name"))
    y = _clean_name_series(merged.get("_yayoi_name"))
    name = z.where(z.str.strip() != "", b)
    name = name.where(name.str.strip() != "", y)
    merged["商品名"] = name

    merged["弥生品番"] = _clean_name_series(merged.get("_yayoi_raw_sku"))
    merged["要確認"] = ""
    # 合計は本社・3F・4F・4F(PKG不備)・物流のみ（弥生は含めない）
    merged["合計"] = merged[PHYSICAL_LOCATION_COLUMNS].sum(axis=1).astype("int64")

    main_table = merged[SHEET1_COLUMNS].copy()

    # 弥生のみに存在し、かつ変換ルールが当てはまらなかった品番 → 独立行として追加
    if not yayoi_needs_check.empty:
        nc = yayoi_needs_check.copy()
        nc["品番"] = NO_SKU_MARK
        nc["要確認"] = NEEDS_CHECK_MARK
        for col in PHYSICAL_LOCATION_COLUMNS:
            nc[col] = 0
        nc["合計"] = 0   # 弥生専用行は本社・3F・4F・4F(PKG不備)・物流が全て0のため合計も0
        nc = nc[SHEET1_COLUMNS]
        main_table = pd.concat([main_table, nc], ignore_index=True)

    main_table["_sort_sku"] = main_table["品番"].replace(NO_SKU_MARK, " ")
    table = (
        main_table
        .sort_values(["_sort_sku", "弥生品番"], kind="stable")
        .drop(columns="_sort_sku")
        .reset_index(drop=True)
    )

    result.table = table
    result.sku_count = int(len(table))
    result.location_totals = {c: int(table[c].sum()) for c in LOCATION_COLUMNS}
    result.grand_total = int(table["合計"].sum())

    # --- シート②: 弥生ゼロ在庫あぶり出しリスト ---
    eligible = table[table["要確認"] != NEEDS_CHECK_MARK]
    abrid = eligible[(eligible["弥生"] == 0) & (eligible["4F"] >= 1)].copy()
    abrid_table = abrid[["品番", "弥生品番", "商品名", "弥生", "4F", "4F(PKG不備)"]].rename(
        columns={"弥生": "弥生在庫数", "4F": "4F在庫数", "4F(PKG不備)": "4F(PKG不備)在庫数"}
    ).reset_index(drop=True)
    result.abrid_table = abrid_table
    result.abrid_count = int(len(abrid_table))

    if result.zaico_missing_sku_rows:
        result.warnings.append(
            f"ZAICO CSV に「型番」が空の行が {result.zaico_missing_sku_rows} 件あり、"
            f"合計 {result.zaico_missing_sku_qty} 個を集計から除外しました。"
        )
    if result.skipped_locations:
        detail = "、".join(f"{k}（{v}行）" for k, v in result.skipped_locations.items())
        known = "、".join(ZAICO_LOCATION_MAP.keys())
        result.warnings.append(
            f"ZAICO CSV に、集計対象の保管場所（{known}）"
            f"以外の保管場所が含まれていました: {detail}。これらは集計していません。"
        )

    try:
        result.output_bytes = _write_excel(result)
        result.output_name = (
            f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        result.success = True
    except Exception as e:  # noqa: BLE001
        result.error = f"Excelの作成に失敗しました。\n{e}"
        result.success = False
    return result


# ====================================================================
# Excel 出力
# ====================================================================

def _style_header_row(ws, row: int, columns: list, header_fill, header_font, border, center):
    for c, name in enumerate(columns, start=1):
        cell = ws.cell(row=row, column=c, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center


def _write_sheet1(wb, result: BuildResult, styles) -> None:
    table = result.table
    assert table is not None
    (thin, border, header_fill, header_font, total_fill, total_font,
     num_fmt, center, left, needs_check_fill) = styles

    ws = wb.active
    ws.title = SHEET1_NAME

    stamp = datetime.now().strftime("%Y/%m/%d %H:%M")
    ws.cell(row=1, column=1, value=f"社内全在庫（拠点別）  作成日時: {stamp}")
    ws.cell(row=1, column=1).font = Font(bold=True)

    header_row = 2
    _style_header_row(ws, header_row, SHEET1_COLUMNS, header_fill, header_font, border, center)

    text_cols = {"品番", "弥生品番", "要確認", "商品名"}
    r = header_row + 1
    for _, row in table.iterrows():
        is_needs_check = row["要確認"] == NEEDS_CHECK_MARK
        for c, name in enumerate(SHEET1_COLUMNS, start=1):
            cell = ws.cell(row=r, column=c, value=row[name])
            cell.border = border
            if name in text_cols:
                cell.alignment = left
            else:
                cell.alignment = center
                cell.number_format = num_fmt
            if is_needs_check:
                cell.fill = needs_check_fill
        r += 1

    total_r = r
    ws.cell(row=total_r, column=1, value="合計")
    ws.cell(row=total_r, column=4, value=f"（{result.sku_count} 品番）")
    for c, name in enumerate(SHEET1_COLUMNS, start=1):
        cell = ws.cell(row=total_r, column=c)
        cell.fill = total_fill
        cell.font = total_font
        cell.border = border
        if name in LOCATION_COLUMNS + ["合計"]:
            col_letter = get_column_letter(c)
            cell.value = f"=SUM({col_letter}{header_row + 1}:{col_letter}{total_r - 1})"
            cell.number_format = num_fmt
            cell.alignment = center
        elif name == "品番":
            cell.alignment = left

    widths = {
        "品番": 16, "弥生品番": 18, "要確認": 10, "商品名": 46,
        "本社": 10, "3F": 10, "4F": 10, "4F(PKG不備)": 13, "物流": 10, "弥生": 10, "合計": 10,
    }
    for c, name in enumerate(SHEET1_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = widths[name]
    ws.freeze_panes = ws.cell(row=header_row + 1, column=5)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(SHEET1_COLUMNS))}{total_r - 1}"


def _write_sheet2(wb, result: BuildResult, styles) -> None:
    abrid_table = result.abrid_table
    assert abrid_table is not None
    (thin, border, header_fill, header_font, total_fill, total_font,
     num_fmt, center, left, needs_check_fill) = styles

    ws = wb.create_sheet(SHEET2_NAME)

    stamp = datetime.now().strftime("%Y/%m/%d %H:%M")
    ws.cell(
        row=1, column=1,
        value=f"弥生ゼロ在庫あぶり出しリスト（弥生在庫0かつ4F在庫1以上）  "
              f"対象 {result.abrid_count} 件  作成日時: {stamp}",
    )
    ws.cell(row=1, column=1).font = Font(bold=True)

    header_row = 2
    _style_header_row(ws, header_row, SHEET2_COLUMNS, header_fill, header_font, border, center)

    text_cols = {"品番", "弥生品番", "商品名"}
    r = header_row + 1
    for _, row in abrid_table.iterrows():
        for c, name in enumerate(SHEET2_COLUMNS, start=1):
            cell = ws.cell(row=r, column=c, value=row[name])
            cell.border = border
            if name in text_cols:
                cell.alignment = left
            else:
                cell.alignment = center
                cell.number_format = num_fmt
        r += 1

    widths = {
        "品番": 16, "弥生品番": 18, "商品名": 46,
        "弥生在庫数": 12, "4F在庫数": 12, "4F(PKG不備)在庫数": 16,
    }
    for c, name in enumerate(SHEET2_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = widths[name]
    ws.freeze_panes = ws.cell(row=header_row + 1, column=4)
    if r > header_row + 1:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(SHEET2_COLUMNS))}{r - 1}"


def _write_excel(result: BuildResult) -> bytes:
    wb = Workbook()

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    total_font = Font(bold=True)
    needs_check_fill = PatternFill("solid", fgColor="FFF2CC")
    num_fmt = "#,##0"
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    styles = (thin, border, header_fill, header_font, total_fill, total_font,
              num_fmt, center, left, needs_check_fill)

    _write_sheet1(wb, result, styles)
    _write_sheet2(wb, result, styles)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
