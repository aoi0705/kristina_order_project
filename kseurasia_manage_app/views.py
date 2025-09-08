from django.shortcuts import render, redirect, get_object_or_404
from django.views.generic import TemplateView
from django.http import HttpResponse
from .forms import OrderExcelUploadForm
from django.contrib import messages

from kseurasia_manage_app.models import *
from django.views.decorators.http import require_POST
import io
import pandas as pd
from typing import Any, Dict, List, Tuple
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Q, Count

import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.worksheet.cell_range import CellRange
import json
from django.conf import settings
import logging
logger = logging.getLogger(__name__)
from django.http import FileResponse
from pathlib import Path
from urllib.parse import quote
from django.utils.encoding import escape_uri_path

from datetime import datetime, date, timedelta
from django.utils import timezone

import re
from collections import OrderedDict,defaultdict

# Create your views here.
def index(request):
    # form = OrderExcelUploadForm()
    # return render(request, 'kseurasia_manage_app/index.html', {'form': form})
    return render(request, 'kseurasia_manage_app/index.html', {'message': 'OK'})

INVOICE_ALIASES: Dict[str, str] = {
  "FLOUVEIL": 19,
  "リレント通常注文": 20,
  "C'BON": 21,
  "Q'1st-1": 22,
  "HIMELABO": 23,
  "SUNSORIT": 24,
  "ELEGADOLL": 25,
  "MAYURI": 26,
  "ATMORE": 27,
  "DIME HEALTH CARE": 28,
  "LAPIDEM": 29,
  "ROSY DROP": 30,
  "ESTLABO": 31,
  "MEROS": 32,
  "COSMEPRO": 33,
  "AFURA": 34,
  "PECLIA": 35,
  "LEJEU": 36,
  "AISHODO": 37,
  "Dr.MEDION": 38,
  "McCoy": 39,
  "Luxces": 40,
  "Evliss": 41,
  "Esthe Pro Labo": 42,
  "COCOCHI　発注書": 43,
  "PURE BIO": 44,
  "BEAUTY GARAGE": 45,
  "DIAMANTE": 46
}

SALES_TABLE_RC_MAP: Dict[str, str] = {
  '(FLOUVEIL)': 'FLOUVEIL',
  '"(RELENT)': 'リレント通常注文',
  "C'BON": "C'BON",
  'Q1st': "Q'1st-1",
  'CHANSON': 'CHANSON',
  'HIMELABO': 'HIMELABO',
  'SUNSORIT': 'SUNSORIT',
  'KYOTOMO': 'KYOTOMO',
  'COREIN': '',
  'ELEGADOLL': 'ELEGADOLL',
  'MAYURI': 'MAYURI',
  'ATMORE': 'ATMORE',
  'OLUPONO': '',
  'DIME HEALTH CARE': 'DIME HEALTH CARE',
  'EMU': 'EMU',
  'CHIKUHODO': 'CHIKUHODO',
  'LAPIDEM': 'LAPIDEM',
  'MARY PLATINUE': 'MARY.P',
  'POD(ROSY DROP)': 'ROSY DROP',
  'CBS(ESTLABO)': 'ESTLABO',
  'DOSHISHA': '',
  'ISTYLE': 'ISTYLE',
  'MEROS': 'MEROS',
  'STAR LAB': 'STARLAB',
  'Beauty Conexion': 'Beauty Conexion',
  'COSMEPRO': 'COSMEPRO',
  'AFURA': 'AFURA',
  'PECLIA': 'PECLIA',
  'OSATO': 'OSATO',
  'HANAKO': 'HANAKO',
  'LEJEU': 'LEJEU',
  'AISHODO': 'AISHODO',
  'CARING JAPAN (RUHAKU)': 'RUHAKU',
  'MEDION': 'Dr.MEDION',
  'McCoy': 'McCoy',
  'URESHINO': 'URESHINO',
  'Luxces': 'Luxces',
  'Evliss': 'Evliss',
  'Pro Labo': 'Esthe Pro Labo',
  'Rey Beaty': 'Rey Beauty',
  'Pure Bio': 'PURE BIO',
  'Diaasjapan': 'Diaas',
  'DIAMANTE': 'DIAMANTE',
  'FAJ': 'FAJ',
  'ＣＨＡＮＳＯＮ': 'CHANSON',
  'Kyo Tomo': 'KYOTOMO',
  'Rey.': 'Rey Beauty',
  '（FLOUVEIL）': 'FLOUVEIL',
  '（RELENT）': 'リレント通常注文',
  '(CBON)': "C'BON",
  '(Q1st)': "Q'1st-1",
  '(姫ラボ）': 'HIMELABO',
  '(SUNSORIT)': 'SUNSORIT',
  'AISEN': '',
  'MARY PL.': 'MARY.P',
  'BEAUTY CONEXION': '',
  'Rey': 'Rey Beauty',
  'FLOUVEIL→\nセンコン': 'FLOUVEIL',
  'センコン→\nKS\n(FLOUVEIL分）': 'FLOUVEIL',
  'RELENT→\nKS': 'リレント通常注文',
  'CBON→\nセンコン': "C'BON",
  "センコン→\nKS\n(C'BON分）": "C'BON",
}
#excelインポート
HEADER_MAP: Dict[str, str] = {
    # ---- 代表列名（左：Excel側正規化名 → 右：OrderContentのフィールド） ----
    "Jan code": "Jan_code",
    "Product Number": "Product_number",
    "Brand name": "Brand_name",

    "SKU Number (RU)": "SKU_number",

    "Produt Name": "Product_name",
    "ORDER": "Order",

    "English Name (RU)": "English_name",
    "RASSIAN NAME (RU)": "Rassian_name",

    "Contents": "Contents",
    "Volume": "Volume",
    "Case Q'ty": "Case_qty",
    "LOT": "Lot",

    "Unit price": "Unit_price",
    "Amount": "Amount",
    "Case Volume": "Case_volume",
    "Case Weight": "Case_weight",
    "Case Q'ty": "Case_qty2",
    "TTL Volume": "TTL_volume",
    "TTL Weight": "TTL_weight",
    "商品サイズ": "Product_size",
    "Unit N/W(kg)": "Unit_nw",
    "TTL N/W(kg)": "TTL_nw",
    "Ingredients": "Ingredients",

    "仕入値": "Purchase_price",
    "仕入値合計": "Purchase_amount",
    "利益": "profit",
    "利益率": "profit_rate",

    # ロシア用
    "Реквизиты ДС": "DS_details",
    "Марка (бренд) ДС": "DS_brandname",
    "Производель ДС": "DS_Manufacturer",
}


# ====== Brand名→テンプレシート名のエイリアス（未解決はTODO） ======
BRAND_TO_SHEET_ALIAS: Dict[str, str] = {
    "CBON": "C'BON","CBON mini sample": "C'BON","CBON SAMPLE": "C'BON","CBON　SAMPLE": "C'BON","CBON　TESTER": "C'BON",
    "EST LABO": "ESTLABO", "EST LABO PRO": "ESTLABO", "ESTLABO": "ESTLABO", "ESTLABO PRO TESTER": "ESTLABO", "ESTLABO STAND": "ESTLABO", "ESTLABO TESTER": "ESTLABO", "LABO+": "ESTLABO", "LABO+ PRO": "ESTLABO", "LABO+ PRO TESTER": "ESTLABO", "LABO+ TESTER": "ESTLABO", "MOTHERMO": "ESTLABO",
    "Elega Doll": "ELEGADOLL","ELEGADOLL TESTER": "ELEGADOLL","Elega Doll PRO": "ELEGADOLL",
    "Evliss": "Evliss", "EVLISS TESTER": "Evliss",
    "LUXCES": "Luxces", "Luxces TESTER": "Luxces",
    "Lapidem": "LAPIDEM", "Lapidem PRO": "LAPIDEM","LAPIDEM": "LAPIDEM", "Lapidem TESTER": "LAPIDEM","Lapidem PRO TESTER": "LAPIDEM",
    "Cosmepro": "COSMEPRO", "Cosmepro TESTER": "COSMEPRO", "Cosmepro PRO": "COSMEPRO",
    "AFURA": "AFURA", "AFURA TESTER": "AFURA",
    "Beaty Conexion": "Beauty Conexion", "Beauty Conexion TESTER": "Beauty Conexion",
    "COCOCHI": "COCOCHI　発注書", "COCOCHI TESTER": "COCOCHI　発注書",  # 全角スペースあり
    "Hime Labo": "HIMELABO",
    "Lejeu TESTER": "LEJEU", "LEJEU": "LEJEU",
    "HANAKO": "HANAKO", "HANAKO TESTER": "HANAKO",
    "McCoy": "McCoy", "McCoy PRO": "McCoy", "McCoy mini pouch": "McCoy", "McCoy TESTER": "McCoy", "McCoy PRO TESTER": "McCoy", "McCoy SANPLE": "McCoy",
    "MEDION": "Dr.Medion", "MEDION PRO": "Dr.Medion", "MEDION sample": "Dr.Medion", "MEDION TESTER": "Dr.Medion",
    "Diaasjapan ": "Diaas", "Diaasjapan mini sample": "Diaas", "Diaasjapan TESTER": "Diaas",
    "ROSY DROP": "ROSY DROP", "ROSY DROP SAMPLE": "ROSY DROP", "ROSY DROP TESTER": "ROSY DROP",
    "Relent": "リレント通常注文", "RELENT PRO": "リレント通常注文", "Relent Sample": "リレント無料提供", "Relent TESTER": "リレント無料提供",
    "AISHODO": "AISHODO", "AISHODO TESTER": "AISHODO",
    "Ajuste": "Ajuste",
    "Atmore": "ATMORE",
    "CHANSON": "CHANSON", "CHANSON TESTER": "CHANSON",
    "Quality 1st": "Q'1st-1", "Quality 1st SAMPLE": "Q'1st-1", "Quality 1st TESTER": "Q'1st-1",
    "Sunsorit": "SUNSORIT", "Sunsorit SAMPLE": "SUNSORIT",
    "DIAMANTE": "DIAMANTE", "DIAMANTE TESTER": "DIAMANTE",
    "Chikuhodo": "CHIKUHODO",
    "SUNTREG": "SUNTREG",
    "MAYURI": "MAYURI", "MAYURI TESTER": "MAYURI",
    "Rey Beauty Studio.": "Rey Beauty",
    "URESHINO": "URESHINO", "URESHINO TESTER": "URESHINO",
    "Kyo Tomo": "KYOTOMO", "Kyo Tomo PRO": "KYOTOMO",
    "Esthe Pro Labo": "Esthe Pro Labo", "Esthe Pro Labo TESTER": "Esthe Pro Labo", "Esthe Pro Labo SAMPLE": "Esthe Pro Labo",
    "MARY PLATINUE": "MARY.P", "MARY PLATINUE TESTER": "MARY.P", "Skin Innovation": "MARY.P", "Skin Innovation PRO": "MARY.P",
    "MEROS": "MEROS", "MEROS TESTER": "MEROS",
    "RUHAKU": "RUHAKU", "RUHAKU TESTER": "RUHAKU", "RUHAKU SAMPLE": "RUHAKU",
    "Olupono": "OLUPONO",
    "Salon de Flouveil": "FLOUVEIL", "Salon de Flouveil SAMPLE": "FLOUVEIL",
    "PECLIA": "PECLIA",
    "Star Lab Cosmetics": "STARLAB",
    "OSATO": "OSATO",
    "Emu No Shizuku": "EMU",
    "BEAUTY GARAGE": "BEAUTY GARAGE",
    "BELEGA": "BELEGA",
    "DENBA": "DENBA",
    "Healing Relax": "Healing Relax",
    "PureBio": "PURE BIO", "Purebio TESTER": "PURE BIO",
    "Dime Health Care PRO": "DIME HEALTH CARE",
    "Lishan": "ISTYLE", "Lishan TESTER": "ISTYLE",
    "Tilla Caps": "FAJ",

    # === TODO: 未解決ブランドの寄せ先が決まり次第ここに追記してください ===
    # "beautygarage": "Rey Beauty",           # 例
    # "dimehealthcarepro": "（対象シート名）",
}

PURCHASE_ORDER_SHEET_ALIAS: Dict[str, str] = {
    #"INV No.": "",
    "Jan code": "Jan code",
    "Brand name": "Brand name",
    "Description of goods": "Produt Name",
    "Case Q'ty": "Case Q'ty",
    "LOT": "LOT",
    "Q'ty": "ORDER",
    "仕入値": "仕入値",
    "仕入値合計": "仕入値合計",
    "ケース容積": "Case Volume",
    "ケース重量": "Case Weight",
    "ケース数量": "Case Q'ty",
    "合計容積": "商品サイズ",
    "合計重量": "Unit N/W(kg)",
    "Unit N/W(kg)": "Unit N/W(kg)",
    "Total N/W(kg)": "TTL N/W(kg)",
    "成分": "Contents",
    "商品名": "Produt Name",
    "JANコード": "Jan code",
    "数量": "ORDER",
    "単位": "Case Q'ty",
    "金額": "Unit price",
    "合計": "Amount",
    "c/s": "ORDER",
    "商品コード": "Product Number",
    "商品名（日本語）": "Produt Name",
    "サイズ": "商品サイズ",
    "ネット重量 中身+容器 （ｇ）": "Unit N/W(kg)",
    "定価": "Unit price",
    "オーダー": "ORDER",
    "仕入合計": "仕入値合計",
    "НАМИМЕНОВАНИЕ": "RASSIAN NAME (RU)",
    "JAN": "Jan code"
}

HEADERS_JSON_PATH = settings.BASE_DIR / "data" / "headers_no_heuristics.json"
NAMES_JSON_PATH = settings.BASE_DIR / "data" / "desc2order_by_kind_strong.json"
PURCHASE_TEMPLATE_PATH = settings.BASE_DIR / "data" / "PurchaseOrder_Template.xlsx"
INVOICE_TEMPLATE_PATH = settings.BASE_DIR / "data" / "InvoicePacking_Template.xlsx"
TEMP_PATH = settings.BASE_DIR / "data" / "tmp"
SALES_TABLE_TEMPLATE_PATH = settings.BASE_DIR / "data" / "SalesTable_Template.xlsx"

INVOICE1 = ["FLOUVEIL","リレント通常注文","C'BON","Q'1st-1"]
INVOICE2 = ["HIMELABO","SUNSORIT","ELEGADOLL","MAYURI","ATMORE","DIME HEALTH CARE","ROSY DROP","LAPIDEM","ESTLABO","MEROS","COSMEPRO","AFURA","PECLIA","LEJEU"]
INVOICE3 = ["Dr.Medion","AISHODO","Luxces","McCoy","Esthe Pro Labo","Evliss","PURE BIO","COCOCHI　発注書","BEAUTY GARAGE","DIAMANTE"]
INVOICE4 = ["リレント通常注文","C'BON","Q'1st-1","CHANSON","LAPIDEM"]
INVOICE5 = ["MARY.P","ROSY DROP","MEROS","B-10","ELEGADOLL""MAYURI","Mediplorer","McCoy","PURE BIO","Glow","ISTYLE"]

# ========= 数量の数値化（空/文字列ゆれ→0） =========
def _as_qty(v: Any) -> float:
    if v is None:
        return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    # ○/●/x 等で発注を表す場合の簡易対応（必要なければ削除してOK）
    if s in {"○", "●", "◯", "x", "X"}:
        return 1.0
    # カンマ等除去
    s = s.replace(",", "")
    try:
        return float(s)
    except:
        return 0.0

# ===== Excel→モデルの列名マッピング（正規化後キー→OrderContentフィールド） =====

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s.lower()

def build_header_index(ws, header_row: int):
    """
    指定したヘッダー行を読み、2つの辞書を返す:
      - exact: {元の文字列そのまま: 列番号}
      - norm : {正規化した文字列: 列番号}
    """
    exact = {}
    normd = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None: 
            continue
        s = str(v).strip()
        if not s:
            continue
        exact[s] = c
        normd[_norm(s)] = c
    return exact, normd

def find_col(header_index_norm: dict, *candidates: str) -> int | None:
    """
    候補名（ゆらぎ含む複数）から最初に見つかった列番号を返す
    """
    for name in candidates:
        col = header_index_norm.get(_norm(name))
        if col:
            return col
    return None

def safe_insert_blank_rows(ws, idx: int, amount: int = 1):
    """
    行 idx の直前に amount 行を挿入する。
    既存の結合セルは、挿入行にかからないように再構築する（= 挿入行は非結合に保つ）。

    ルール:
      - 完全に idx より上の結合…そのまま
      - 完全に idx より下の結合…行数ぶん下へシフト
      - idx をまたぐ結合（min_row < idx <= max_row）…上側と下側に分割して再結合
        （＝挿入行は結合に含めない）
    """
    # 1) 既存の結合を控える
    old_ranges = list(ws.merged_cells.ranges)

    # 2) いったん全部ほどく
    for rng in old_ranges:
        ws.unmerge_cells(str(rng))

    # 3) 行を挿入
    ws.insert_rows(idx, amount=amount)

    # 4) 結合を再構築
    for rng in old_ranges:
        r1, r2, c1, c2 = rng.min_row, rng.max_row, rng.min_col, rng.max_col

        if r2 < idx:
            # そのまま
            new_parts = [(r1, r2, c1, c2)]
        elif r1 >= idx:
            # まとめて下へシフト
            new_parts = [(r1 + amount, r2 + amount, c1, c2)]
        else:
            # idx をまたいでいる → 上下に分割（挿入行は含めない）
            upper = (r1, idx - 1, c1, c2) if r1 <= idx - 1 else None
            lower = (idx + amount, r2 + amount, c1, c2) if idx <= r2 else None
            new_parts = [p for p in (upper, lower) if p is not None]

        # 面積>1 のものだけ再結合（1セルは結合不要）
        for nr1, nr2, nc1, nc2 in new_parts:
            if nr1 < nr2 or nc1 < nc2:
                ws.merge_cells(start_row=nr1, start_column=nc1, end_row=nr2, end_column=nc2)

def safe_insert_blank_cols(ws, idx: int, amount: int = 1):
    """
    列 idx の直前に amount 列を挿入する。
    既存の結合セルは、挿入列にかからないように再構築する（= 挿入列は非結合に保つ）。

    ルール:
      - 完全に idx より左の結合 … そのまま
      - 完全に idx 以右の結合 … 列数ぶん右へシフト
      - idx をまたぐ結合（min_col < idx <= max_col）… 左右に分割して再結合
        （＝挿入列は結合に含めない）
    """
    # 1) 現在の結合を退避
    old_ranges = list(ws.merged_cells.ranges)

    # 2) いったん全て解除
    for rng in old_ranges:
        ws.unmerge_cells(str(rng))

    # 3) 列を挿入
    ws.insert_cols(idx, amount=amount)

    # 4) 結合を再構築
    for rng in old_ranges:
        r1, r2, c1, c2 = rng.min_row, rng.max_row, rng.min_col, rng.max_col

        if c2 < idx:
            # 挿入位置より左側に完全に存在 → そのまま
            new_parts = [(r1, r2, c1, c2)]
        elif c1 >= idx:
            # 挿入位置より右側に完全に存在 → 右へシフト
            new_parts = [(r1, r2, c1 + amount, c2 + amount)]
        else:
            # 挿入列をまたぐ → 左右に分割（挿入列は含めない）
            left  = (r1, r2, c1, idx - 1)              if c1 <= idx - 1 else None
            right = (r1, r2, idx + amount, c2 + amount) if idx <= c2   else None
            new_parts = [p for p in (left, right) if p is not None]

        # 面積が 1×1 を超えるときだけ再結合（1セルは結合不要）
        for nr1, nr2, nc1, nc2 in new_parts:
            if (nr2 - nr1) >= 1 or (nc2 - nc1) >= 1:
                ws.merge_cells(start_row=nr1, start_column=nc1, end_row=nr2, end_column=nc2)

# ========= 取り込み =========
def import_orders(request):
    """
    .xlsx アップロードを受け取り、OrderContent に一括登録。
    ・ヘッダー自動推定（必要なら固定 header=2 でOK）
    ・HEADER_MAP で列マッピング
    ・Order（数量）が > 0 の行だけ登録
    """
    if request.method == "POST":
        form = OrderExcelUploadForm(request.POST, request.FILES)
        if not form.is_valid():
            messages.error(request, "フォームの入力に誤りがあります。")
            return render(request, "kseurasia_manage_app/import_orders.html", {"form": form})

        f = form.cleaned_data["file"]
        sheet_name = form.cleaned_data.get("sheet_name") or 0  # 未指定→先頭シート

        # dest = form.cleaned_data.get("destination") or ""
        # client = form.cleaned_data.get("client") or ""

        # BytesIO 2回読むのでseek可の別コピーを用意
        raw = io.BytesIO(f.read())
        xl_raw = f.read()
        in_buf_for_xl = io.BytesIO(xl_raw) 

        hdr = 2

        # top_df = pd.read_excel(io.BytesIO(raw.getvalue()), sheet_name=sheet_name, engine="openpyxl",
        #                     header=None, nrows=3, usecols="A:D")
        # a1 = top_df.iat[0, 0]
        # d3 = top_df.iat[2, 3]        

        # if "ROYAL COSMETICS FOR DUBAI" in a1:
        #     hdr = 4
        # elif "ROYAL COSMETICS" in a1:
        #     hdr = 2
        # elif "KSユーラシア株式会社様" in d3:
        #     hdr = 19
        # 2) 本読み込み
        try:
            df = pd.read_excel(io.BytesIO(raw.getvalue()), sheet_name=sheet_name,
                               engine="openpyxl", header=hdr)
        except Exception as e:
            messages.error(request, f"Excelの読み込みに失敗しました: {e}")
            return render(request, "kseurasia_manage_app/import_orders.html", {"form": form})

        if df.empty:
            messages.warning(request, "Excelにデータがありません。")
            return render(request, "kseurasia_manage_app/import_orders.html", {"form": form})

        # 3) 列マッピング
        original_cols = [str(c) for c in df.columns]
        normalized_cols = original_cols

        mapped: List[Tuple[str, str]] = []   # (Excel列名, モデル列名)
        unmapped: List[str] = []             # 参考表示用
        for oc, nc in zip(original_cols, normalized_cols):
            if nc in HEADER_MAP:
                mapped.append((oc, HEADER_MAP[nc]))
            else:
                unmapped.append(oc)

        if not mapped:
            messages.error(request, "モデルに対応できる列が見つかりませんでした。ヘッダー名をご確認ください。")
            return render(request, "kseurasia_manage_app/import_orders.html",
                          {"form": form, "columns": original_cols})

        # 4) バッチ作成（紐付け）
        batch = ImportBatch.objects.create(
            source_filename=(getattr(f, "name", "") or "")[:255],
            sheet_name=str(sheet_name) if sheet_name != 0 else "",
        )

        # 5) 取り込み（ORDER>0 の行だけ）
        objects: List[OrderContent] = []
        PurchaseOrder_list = []
        for _, row in df.iterrows():
            data: Dict[str, Any] = {}
            for excel_col, model_field in mapped:
                # 値取り出し
                val = row.get(excel_col)
                # 文字列化ナシ：数値列は数値のまま受けてOK、モデルはCharFieldなので下でstr化
                data[model_field] = None if pd.isna(val) else val

            # 数量チェック（Orderがない場合は非発注としてスキップ）
            qty = _as_qty(data.get("Order"))
            if qty <= 0:
                continue  # ←★ここが「数量が入っている行だけ」の肝

            # ブランド名取得
            brand_dict = {}
            brand_raw = row.get("Brand name")
            #logger.info(f"brand_raw={brand_raw}")
            if brand_raw in BRAND_TO_SHEET_ALIAS:
                sheet_name = BRAND_TO_SHEET_ALIAS[brand_raw]
            else:
                logger.warning(f"brand_rawが見つかりません continue")
                continue
            brand_dict[sheet_name] = row
            PurchaseOrder_list.append(brand_dict)

            # 文字列化（CharField前提）
            for k, v in list(data.items()):
                if v is None:
                    continue
                data[k] = str(v)

            data["batch"] = batch
            #data["supplier"] = "ROYAL COOSMETICS"
            objects.append(OrderContent(**data))

        if not objects:
            messages.warning(request, "ORDER（数量）が入っている行が見つかりませんでした。")
            return render(request, "kseurasia_manage_app/import_orders.html", {"form": form})

        # 6) 一括登録
        try:
            with transaction.atomic():
                OrderContent.objects.bulk_create(objects, batch_size=500)
        except Exception as e:
            messages.error(request, f"登録中にエラーが発生しました: {e}")
            return render(request, "kseurasia_manage_app/import_orders.html", {"form": form})
        
        #7 発注書作成
        with HEADERS_JSON_PATH.open("r", encoding="utf-8") as f:
            headers_summary = json.load(f)  # dict として使えます
        with NAMES_JSON_PATH.open("r", encoding="utf-8") as f:
            names_summary = json.load(f)  # dict として使えます

        insert_dict = {}

        thin_border = Border(
            left=Side(style="thin"),   # 左
            right=Side(style="thin"),  # 右
            top=Side(style="thin"),    # 上
            bottom=Side(style="thin")  # 下
        )
        purchase_font = Font(
            name="Arial",   # フォント名（例: メイリオ）
            size=8,         # 文字サイズ
            bold=False,       # 太字
            italic=False,    # 斜体
            color="000000"   # 赤色（RGB指定）
        )
        tester_list = ["TESTER", "SAMPLE", "sample", "mini sample", "SAMPLE", "SAMPLE"]
        wb = openpyxl.load_workbook(PURCHASE_TEMPLATE_PATH, data_only=False)
        for brand_dict in PurchaseOrder_list:
            sheet_name = next(iter(brand_dict))   
            #logger.info(f"sheet_name={sheet_name}")
            if sheet_name not in brand_dict:
                logger.warning(f"sheet_name={sheet_name} not in brand_dict continue")
                continue
            data_row = brand_dict[sheet_name]
            order_ProductName = data_row.get("Produt Name")
            tester_flg = False
            if any(tester in order_ProductName for tester in tester_list):
                logger.info(f"order_ProductName={order_ProductName} is tester continue")
                tester_flg = True

            #wb[str(brand_dict.keys()[0])].cell(row=1, column=1).value = "test"
            if sheet_name not in headers_summary:
                logger.warning(f"sheet_name={sheet_name} not in headers_summary continue")
                continue
            sheet_config = headers_summary[sheet_name]
            ws = wb[sheet_name]

            first_header_row = sheet_config.get("first_header")
            first_now_row = first_header_row + 1
            second_header_row = sheet_config.get("second_header")
            second_now_row = second_header_row + 1

            if sheet_name in insert_dict:
                add_row = insert_dict[sheet_name]
            elif sheet_name not in insert_dict:
                add_row = 0
            
            #ws.insert_rows(first_now_row)
            if tester_flg == False:
                safe_insert_blank_rows(ws, idx=first_now_row, amount=1)
                add_row += 1
            elif tester_flg == True:
                second_now_row = second_now_row + add_row
                safe_insert_blank_rows(ws, idx=second_now_row, amount=1)

            for col in range(1,20):
                if tester_flg == False:
                    if ws.cell(first_header_row,col).value is not None:
                        ws.cell(first_now_row,col).border = thin_border
                        ws.cell(first_now_row,col).font = purchase_font

                    if ws.cell(first_header_row,col).value in PURCHASE_ORDER_SHEET_ALIAS:
                        PurchaseCell_value = data_row.get(PURCHASE_ORDER_SHEET_ALIAS[ws.cell(first_header_row,col).value])
                        if PurchaseCell_value is None or PurchaseCell_value == "Nan" or PurchaseCell_value == "nan":
                            PurchaseCell_value = ""
                        ws.cell(first_now_row,col).value = str(PurchaseCell_value)
                elif tester_flg == True:
                    if ws.cell(second_header_row,col).value is not None:
                        ws.cell(second_now_row,col).border = thin_border
                        ws.cell(second_now_row,col).font = purchase_font

                    if ws.cell(second_header_row,col).value in PURCHASE_ORDER_SHEET_ALIAS:
                        PurchaseCell_value = data_row.get(PURCHASE_ORDER_SHEET_ALIAS[ws.cell(second_header_row,col).value])
                        if PurchaseCell_value is None or PurchaseCell_value == "Nan" or PurchaseCell_value == "nan":
                            PurchaseCell_value = ""
                        ws.cell(second_now_row,col).value = str(PurchaseCell_value)
            insert_dict[sheet_name] = add_row

        # 現在日時を文字列化
        now_str = datetime.now().strftime("%Y%m%d%H%M")
        save_path = Path(TEMP_PATH) / f"{now_str}_{batch.id}_PurchaseOrder.xlsx"
        wb.save(save_path)

        batch.PurchaseOrder_file = str(save_path)              # 例: "C:\...\tmp\20250905_36_PurchaseOrder.xlsx"
        # batch.PurchaseOrder_file = save_path.name            # 例: "20250905_36_PurchaseOrder.xlsx"（名前だけ保存したい場合）
        batch.save(update_fields=["PurchaseOrder_file"])

        try:
            wb.close()
        except AttributeError:
            pass

        #インボイス・パッキングリスト作成
        invoice_font = Font(
            name="Times New Roman",   # フォント名（例: メイリオ）
            size=14,         # 文字サイズ
            bold=False,       # 太字
            italic=False,    # 斜体
            color="000000"   # 赤色（RGB指定）
        )
        packing_font = Font(
            name="Times New Roman",   # フォント名（例: メイリオ）
            size=8,         # 文字サイズ
            bold=False,       # 太字
            italic=False,    # 斜体
            color="000000"   # 赤色（RGB指定）
        )

        invoice_wb = openpyxl.load_workbook(INVOICE_TEMPLATE_PATH, data_only=False)
        invoice_row_dict = {
            "Invoice1":14,
            "Invoice2":14,
            "Invoice3":14,
            "Invoice4":14,
            "Invoice5":14,
            "PL1":13,
            "PL2":14,
            "PL3":13,
        }
        price_dict = {}
        qty_dict = {}
        
        for brand_dict in PurchaseOrder_list:
            sheet_name = next(iter(brand_dict))   
            #logger.info(f"sheet_name={sheet_name}")
            if sheet_name not in brand_dict:
                logger.warning(f"sheet_name={sheet_name} not in brand_dict continue")
                continue
            data_row = brand_dict[sheet_name]
            order_ProductName = data_row.get("Produt Name")

            if  any(tester in order_ProductName for tester in tester_list):
                if sheet_name in INVOICE4:
                    invoice_ws = invoice_wb["Invoice-4 (TESTER)"]
                    invoice_row = invoice_row_dict["Invoice4"]
                elif sheet_name in INVOICE5:
                    invoice_ws = invoice_wb["Invoice-5 (TESTER) "]
                    invoice_row = invoice_row_dict["Invoice5"]
            elif sheet_name in INVOICE1:
                invoice_ws = invoice_wb["Invoice-1 "]
                packing_ws = invoice_wb["PL1"]
                invoice_row = invoice_row_dict["Invoice1"]
                packing_row = invoice_row_dict["PL1"]
            elif sheet_name in INVOICE2:
                invoice_ws = invoice_wb["Invoice-2"]
                packing_ws = invoice_wb["PL2"]
                invoice_row = invoice_row_dict["Invoice2"]
                packing_row = invoice_row_dict["PL2"]
            elif sheet_name in INVOICE3:
                invoice_ws = invoice_wb["Invoice-3"]
                packing_ws = invoice_wb["PL3"]
                invoice_row = invoice_row_dict["Invoice3"]
                packing_row = invoice_row_dict["PL3"]
            else:
                logger.warning(f"sheet_name={sheet_name} not in INVOICE continue")
                continue
                
            safe_insert_blank_rows(invoice_ws, idx=invoice_row, amount=1)
            invoice_ws.cell(invoice_row,2).value = data_row.get("Produt Name")
            invoice_ws.cell(invoice_row,3).value = data_row.get("English Name (RU)")
            invoice_ws.cell(invoice_row,4).value = data_row.get("RASSIAN NAME (RU)")
            invoice_ws.cell(invoice_row,5).value = data_row.get("ORDER")
            invoice_ws.cell(invoice_row,6).value = f"¥{data_row.get("Unit price"):,}"
            invoice_ws.cell(invoice_row,7).value = f"¥{int(data_row.get("ORDER")) * int(data_row.get("Unit price")):,}"
            invoice_ws.cell(invoice_row,8).value = data_row.get("Unit N/W(kg)")
            invoice_ws.cell(invoice_row,9).value = data_row.get("TTL N/W(kg)")
            invoice_ws.cell(invoice_row,10).value = data_row.get("SKU Number (RU)")
            invoice_ws.cell(invoice_row,11).value = data_row.get("Реквизиты ДС")
            invoice_ws.cell(invoice_row,12).value = data_row.get("Марка (бренд) ДС")
            invoice_ws.cell(invoice_row,13).value = data_row.get("Производель ДС")
            for col in range(2, 14):
                invoice_ws.cell(invoice_row, col).font = invoice_font

            safe_insert_blank_rows(packing_ws, idx=packing_row, amount=1)
            packing_ws.cell(packing_row,3).value = data_row.get("Case Weight")
            packing_ws.cell(packing_row,4).value = data_row.get("ORDER")
            packing_ws.cell(packing_row,5).value = data_row.get("Produt Name")
            packing_ws.cell(packing_row,6).value = data_row.get("English Name (RU)")
            packing_ws.cell(packing_row,7).value = data_row.get("RASSIAN NAME (RU)")
            packing_ws.cell(packing_row,8).value = data_row.get("Contents")
            packing_ws.cell(packing_row,9).value = data_row.get("Case Volume")
            packing_ws.cell(packing_row,10).value = data_row.get("TTL Volume")
            packing_ws.cell(packing_row,11).value = data_row.get("Unit N/W(kg)")
            packing_ws.cell(packing_row,14).value = data_row.get("Case Q'ty")
            packing_ws.cell(packing_row,15).value = data_row.get("TTL Weight")
            packing_ws.cell(packing_row,16).value = data_row.get("Unit price")
            packing_ws.cell(packing_row,17).value = int(data_row.get("ORDER")) * int(data_row.get("Unit price"))
            packing_ws.cell(packing_row,18).value = data_row.get("SKU Number (RU)")
            packing_ws.cell(packing_row,19).value = data_row.get("Реквизиты ДС")
            packing_ws.cell(packing_row,20).value = data_row.get("Марка (бренд) ДС")
            packing_ws.cell(packing_row,21).value = data_row.get("Производель ДС")

            #各シートのTotal部分に記載
            # packing_ws.cell(packing_row+1,5).value = int(packing_ws.cell(packing_row+1,5).value) + int(data_row.get("ORDER"))
            # packing_ws.cell(packing_row+1,6).value = int(packing_ws.cell(packing_row+1,6).value) + int(data_row.get("Unit price"))
            # packing_ws.cell(packing_row+1,7).value = int(packing_ws.cell(packing_row+1,7).value) + (int(data_row.get("ORDER")) * int(data_row.get("Unit price")))
            # packing_ws.cell(packing_row+1,8).value = packing_ws.cell(packing_row+1,8).value + data_row.get("Unit N/W(kg)")
            # packing_ws.cell(packing_row+1,9).value = packing_ws.cell(packing_row+1,9).value + data_row.get("TTL N/W(kg)")

            for col in range(3, 21):
                packing_ws.cell(packing_row, col).font = packing_font

            if sheet_name not in price_dict:
                price_dict[sheet_name] = int(data_row.get("ORDER")) * int(data_row.get("Unit price"))
                qty_dict[sheet_name] = int(data_row.get("ORDER"))
            else:
                price_dict[sheet_name] += int(data_row.get("ORDER")) * int(data_row.get("Unit price"))
                qty_dict[sheet_name] += int(data_row.get("ORDER"))

            if  any(tester in order_ProductName for tester in tester_list):
                if sheet_name in INVOICE4:
                    invoice_row_dict["Invoice4"] += 1
                elif sheet_name in INVOICE5:
                    invoice_row_dict["Invoice5"] += 1
                else:
                    logger.warning(f"sheet_name={sheet_name} not in INVOICE4/5_TESTER_SAMPLE continue")
                    continue
            elif sheet_name in INVOICE1:
                invoice_row_dict["Invoice1"] += 1
                invoice_row_dict["PL1"] += 1
            elif sheet_name in INVOICE2:
                invoice_row_dict["Invoice2"] += 1
                invoice_row_dict["PL2"] += 1
            elif sheet_name in INVOICE3:
                invoice_row_dict["Invoice3"] += 1
                invoice_row_dict["PL3"] += 1
            else:
                logger.warning(f"sheet_name={sheet_name} not in INVOICE continue")
                continue

        total_qty = 0
        total_price = 0
        invoice_head_ws = invoice_wb["INVOICE(無償サンプル抜き価格)"]
        for invoice_key, invoice_value in price_dict.items():
            if invoice_key in INVOICE_ALIASES:
                invoice_head_ws.cell(row=INVOICE_ALIASES[invoice_key], column=8).value = qty_dict[invoice_key]
                invoice_head_ws.cell(row=INVOICE_ALIASES[invoice_key], column=9).value = f"¥{invoice_value:,}"
                total_qty += qty_dict[invoice_key]
                total_price += invoice_value

        invoice_head_ws.cell(row=48, column=8).value = f"{total_qty:,}"
        invoice_head_ws.cell(row=48, column=9).value = f"¥{total_price:,}"

        # 現在日時を文字列化
        now_str = datetime.now().strftime("%Y%m%d%H%M")
        save_path = Path(TEMP_PATH) / f"{now_str}_{batch.id}_InvoicePackingList.xlsx"
        invoice_wb.save(save_path)

        batch.InvoicePacking_file = str(save_path)          # 例: "C:\...\tmp\20250905_36_PurchaseOrder.xlsx"
        batch.buyers = "ROYAL COOSMETICS"
        # batch.PurchaseOrder_file = save_path.name            # 例: "20250905_36_PurchaseOrder.xlsx"（名前だけ保存したい場合）
        batch.save(update_fields=["InvoicePacking_file","buyers"])

        try:
            invoice_wb.close()
        except AttributeError:
            pass


       # 結果画面に遷移（PRGが良ければ redirect でも可）
        return redirect("import_batch_list")

    # GET: フォーム表示
    form = OrderExcelUploadForm()
    return render(request, "kseurasia_manage_app/index.html", {"form": form})

def order_list(request):
    """
    注文の一覧（初期表示で全件：新しい順）
    - ページネーション
    - 任意でバッチ/キーワード検索も可能（未指定なら全て）
    """
    batch_id = request.GET.get("batch")
    q = request.GET.get("q", "")
    page = request.GET.get("page", 1)

    # ベースクエリ：新しい順
    qs = OrderContent.objects.select_related("batch").order_by("-id")

    # 任意：バッチ絞り込み
    if batch_id:
        qs = qs.filter(batch_id=batch_id)

    # 任意：簡易検索（必要な項目は適宜追加）
    if q:
        qs = qs.filter(
            Q(Jan_code__icontains=q) |
            Q(Product_name__icontains=q) |
            Q(Brand_name__icontains=q) |
            Q(SKU_number__icontains=q)
        )

    # ページネーション
    paginator = Paginator(qs, 50)  # 1ページ50件
    page_obj = paginator.get_page(page)

    # バッチ選択用（ドロップダウン）
    batches = ImportBatch.objects.all().only("id", "created_at", "source_filename")

    return render(request, "kseurasia_manage_app/order_list.html", {
        "page_obj": page_obj,
        "total_count": qs.count(),
        "batches": batches,
        "selected_batch_id": batch_id,
        "q": q,
    })


def order_detail(request, pk: int):
    """1件の全情報を表示"""
    obj = get_object_or_404(OrderContent.objects.select_related("batch"), pk=pk)
    ctx = {"obj": obj}
    # バッチ一覧経由で来た時は戻り先を出す
    from_batch_id = request.GET.get("from_batch")
    if from_batch_id:
        ctx["from_batch_id"] = from_batch_id
    return render(request, "kseurasia_manage_app/order_detail.html", ctx)

def batch_order_list(request, batch_id: int):
    """
    指定バッチに属するオーダー一覧（ページネーション付き）
    """
    batch = get_object_or_404(ImportBatch, pk=batch_id)
    page = request.GET.get("page", 1)
    q = request.GET.get("q", "").strip()

    qs = (
        OrderContent.objects
        .select_related("batch")
        .filter(batch_id=batch.id)
        .order_by("-id")
    )
    if q:
        qs = qs.filter(
            Q(Jan_code__icontains=q) |
            Q(Product_name__icontains=q) |
            Q(Brand_name__icontains=q) |
            Q(SKU_number__icontains=q)
        )

    paginator = Paginator(qs, 50)
    page_obj = paginator.get_page(page)

    ctx = {
        "batch": batch,
        "page_obj": page_obj,
        "total_count": qs.count(),
        "q": q,
    }
    return render(request, "kseurasia_manage_app/batch_order_list.html", ctx)

def batch_order_detail(request, batch_id: int, pk: int):
    """
    指定バッチ配下で、かつ pk が属している場合のみ詳細表示
    """
    batch = get_object_or_404(ImportBatch, pk=batch_id)
    obj = get_object_or_404(
        OrderContent.objects.select_related("batch"),
        pk=pk,
        batch_id=batch.id
    )
    # テンプレは既存の order_detail.html を再利用
    return render(
        request,
        "kseurasia_manage_app/order_detail.html",
        {"obj": obj, "from_batch_id": batch.id}
    )

@require_POST
def delete_import_batch(request, batch_id: int):
    """バッチと関連OrderContentを削除。関連ファイルも削除。"""
    batch = get_object_or_404(ImportBatch, pk=batch_id)

    # 関連ファイル（保存していれば）を削除
    for attr in ("PurchaseOrder_file", "InvoicePacking_file"):
        val = getattr(batch, attr, "") or ""
        if val:
            p = Path(val)
            if not p.is_absolute():
                p = Path(TEMP_PATH) / val
            try:
                if p.exists() and p.is_file():
                    p.unlink()
            except Exception as e:
                logger.warning("Failed to delete file %s: %s", p, e)

    # OrderContent は on_delete=SET_NULL のため手動削除
    OrderContent.objects.filter(batch_id=batch.id).delete()

    # バッチ削除
    batch.delete()

    messages.success(request, "バッチと関連オーダーを削除しました。")
    return redirect("import_batch_list")

#以下はインポートバッチ関連
def import_batch_list(request):
    q = request.GET.get("q", "").strip()
    date_from = request.GET.get("from", "").strip()
    date_to   = request.GET.get("to", "").strip()
    page = int(request.GET.get("page", 1))

    # ここを修正：ordercontent → items（モデルの related_name に合わせる）
    batches = (
        ImportBatch.objects.all()
        .annotate(item_count=Count("items", distinct=True))  # ★ 修正
        .order_by("-id")
    )

    if q:
        batches = batches.filter(
            Q(source_filename__icontains=q) |
            Q(sheet_name__icontains=q) |
            Q(id__icontains=q)
        )

    if date_from:
        batches = batches.filter(created_at__date__gte=date_from)
    if date_to:
        batches = batches.filter(created_at__date__lte=date_to)

    paginator = Paginator(batches, 20)
    page_obj = paginator.get_page(page)

    ctx = {
        "page_obj": page_obj,
        "q": q,
        "date_from": date_from,
        "date_to": date_to,
    }
    return render(request, "kseurasia_manage_app/import_batch_list.html", ctx)


# --- 以降はダミー（あなたが実装） -----------------
from django.http import HttpResponseNotAllowed
from django.http import FileResponse, Http404

def download_purchase_order(request, batch_id: int):
    batch = get_object_or_404(ImportBatch, pk=batch_id)

    filename_or_path = getattr(batch, "PurchaseOrder_file", "") or ""
    if not filename_or_path:
        raise Http404("このバッチには発注書ファイルが登録されていません。")

    p = Path(filename_or_path)
    # 相対（ファイル名のみ）で保存している場合は TEMP_PATH を前置
    if not p.is_absolute():
        p = Path(TEMP_PATH) / filename_or_path

    if not p.exists() or not p.is_file():
        raise Http404("発注書ファイルが見つかりません。")

    f = open(p, "rb")
    resp = FileResponse(
        f,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    quoted = quote(p.name)
    # ダウンロード時のファイル名（UTF-8対応）
    resp["Content-Disposition"] = f'attachment; filename="purchase_order.xlsx"; filename*=UTF-8\'\'{quoted}'
    return resp

def export_invoice_packing(request, batch_id: int):
    batch = get_object_or_404(ImportBatch, pk=batch_id)

    filename_or_path = getattr(batch, "InvoicePacking_file", "") or ""
    if not filename_or_path:
        raise Http404("このバッチにはIVPLファイルが登録されていません。")

    p = Path(filename_or_path)
    # 相対（ファイル名のみ）で保存している場合は TEMP_PATH を前置
    if not p.is_absolute():
        p = Path(TEMP_PATH) / filename_or_path

    if not p.exists() or not p.is_file():
        raise Http404("IVPLファイルファイルが見つかりません。")

    f = open(p, "rb")
    resp = FileResponse(
        f,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    quoted = quote(p.name)
    # ダウンロード時のファイル名（UTF-8対応）
    resp["Content-Disposition"] = f'attachment; filename="InvoicePackingList.xlsx"; filename*=UTF-8\'\'{quoted}'
    return resp

def report_console(request):
    return render(request, "kseurasia_manage_app/reports.html")

#以下売り上げ票関連
def _parse_month_yyyy_mm(s: str) -> date:
    """'YYYY-MM' → その月の1日(naive date)"""
    try:
        dt = datetime.strptime(s.strip(), "%Y-%m")
        return date(dt.year, dt.month, 1)
    except Exception:
        raise Http404("start_month / end_month は 'YYYY-MM' 形式で指定してください。")

def _month_end(d: date) -> date:
    """その月の末日(naive date)を返す"""
    y, m = d.year, d.month
    if m == 12:
        return date(y, 12, 31)
    return date(y, m + 1, 1) - timedelta(days=1)

def _get_report_params_range(request):
    """
    レポート用：クエリから
      - clients: 複数 (client=A&client=B…)
      - start_month: 'YYYY-MM'
      - end_month  : 'YYYY-MM'
      - format     : 'csv' or 'pdf'（デフォルト csv）
    を取り出し、期間 [start,end) と表示用ラベルを返す。
    """
    clients = request.GET.getlist("client")  # 複数
    start_ym = (request.GET.get("start_month") or "").strip()
    end_ym   = (request.GET.get("end_month") or "").strip()
    fmt = (request.GET.get("format") or "csv").lower()

    if not clients or not start_ym or not end_ym:
        raise ValueError("client（1つ以上）, start_month, end_month は必須です")

    start_dt, end_dt, label = _month_range_aware(start_ym, end_ym)
    return clients, start_dt, end_dt, fmt, label

def ym_to_year(ym: str) -> int:
    """'YYYY-MM' から年(YYYY)だけ取り出す。'YYYY' 単体も許容。"""
    ym = (ym or "").strip()
    try:
        return datetime.strptime(ym, "%Y-%m").year
    except ValueError:
        if re.fullmatch(r"\d{4}", ym):
            return int(ym)
        raise ValueError("ym は 'YYYY' または 'YYYY-MM' で指定してください")
    
def orders_imported_in_year(year: int):
    """指定年(YYYY)に ImportBatch.created_at が属する OrderContent を全件返す。"""
    start = timezone.make_aware(datetime(year,     1, 1, 0, 0, 0))
    end   = timezone.make_aware(datetime(year + 1, 1, 1, 0, 0, 0))
    return (
        OrderContent.objects
        .select_related("batch")
        .filter(batch__created_at__gte=start, batch__created_at__lt=end)
        .order_by("id")
    )

def _parse_ym(ym: str) -> tuple[int, int]:
    """'YYYY-MM' → (year, month)。不正は ValueError。"""
    dt = datetime.strptime(ym, "%Y-%m")
    return dt.year, dt.month

def _month_range_aware(start_ym: str, end_ym: str) -> tuple[datetime, datetime, str]:
    """
    [start, end) の aware datetime（終了は『終了月の翌月1日 00:00』未満）。
    併せてファイル名用のラベルも返す（例: '2024-04_2024-06'）。
    """
    sy, sm = _parse_ym(start_ym)
    ey, em = _parse_ym(end_ym)

    start = timezone.make_aware(datetime(sy, sm, 1, 0, 0, 0))
    # end = 終了月の翌月1日
    if em == 12:
        end = timezone.make_aware(datetime(ey + 1, 1, 1, 0, 0, 0))
    else:
        end = timezone.make_aware(datetime(ey, em + 1, 1, 0, 0, 0))

    if start >= end:
        raise ValueError("開始年月は終了年月以前にしてください")

    label = f"{start_ym}_{end_ym}"
    return start, end, label

def import_batches_between(start_dt: datetime, end_dt: datetime):
    """期間に作成された ImportBatch を返す queryset。"""
    return (
        ImportBatch.objects
        .filter(created_at__gte=start_dt, created_at__lt=end_dt)
        .order_by("id")
    )

def orders_between_by_batch_created(start_dt: datetime, end_dt: datetime):
    """
    期間に作成されたバッチに紐づく OrderContent を返す queryset。
    （売上表など“取り込み期間ベース”で集計したい場合）
    """
    return (
        OrderContent.objects
        .select_related("batch")
        .filter(batch__created_at__gte=start_dt, batch__created_at__lt=end_dt)
        .order_by("id")
    )

def _download_response(binary: bytes, filename: str, content_type: str):
    resp = HttpResponse(binary, content_type=content_type)
    # 日本語ファイル名対応
    quoted = escape_uri_path(filename)
    resp["Content-Disposition"] = f'attachment; filename="{quoted}"'
    return resp

# ---- 以下4つを実装すればフロントのDLボタンが動きます ----

def reports_sales_export(request):
    # 例: CSV/PDF生成用のエンドポイント
    client_list = ["ROYAL COOSMETICS"]  # 今回は固定
    clients, start_dt, end_dt, fmt, label = _get_report_params_range(request)
    qs = orders_between_by_batch_created(start_dt, end_dt)

    batches = (
        import_batches_between(start_dt, end_dt)
        .prefetch_related("items")          # OrderContent を一括プリフェッチ
        .order_by("id")
    )

    # Python側で「バッチ → そのオーダー一覧」に整形
    result = OrderedDict()
    for b in batches:
        result[b.id] = {
            "batch": {
                "id": b.id,
                "created_at": b.created_at,
                "source_filename": b.source_filename,
                "sheet_name": b.sheet_name,
                "buyers": b.buyers,
            },
            "orders": [
                {
                    "id": oc.id,
                    "Jan_code": oc.Jan_code,
                    "Product_name": oc.Product_name,
                    "Brand_name": oc.Brand_name,
                    "Purchase_price": oc.Purchase_price,
                    "Purchase_amount": oc.Purchase_amount,
                    "profit": oc.profit,
                    "profit_rate": oc.profit_rate,
                    "Amount": oc.Amount,
                }
                for oc in b.items.all()
            ],
        }

    wb = openpyxl.load_workbook(SALES_TABLE_TEMPLATE_PATH, data_only=False)
    ws = wb["R&C"]  # ★今回の帳票は R&C シート前提なので先に取得

    

def reports_ar_export(request):
    clients, start_dt, end_dt, fmt, label = _get_report_params_range(request)
    qs = orders_between_by_batch_created(start_dt, end_dt)
    filename = f"accounts_receivable_{label}.{fmt}"
    content = b""
    ctype = "text/csv" if fmt == "csv" else "application/pdf"
    return _download_response(content, filename, ctype)

def reports_ap_export(request):
    clients, start_dt, end_dt, fmt, label = _get_report_params_range(request)
    qs = orders_between_by_batch_created(start_dt, end_dt)
    filename = f"accounts_payable_{label}.{fmt}"
    content = b""
    ctype = "text/csv" if fmt == "csv" else "application/pdf"
    return _download_response(content, filename, ctype)

def reports_cashflow_export(request):
    clients, start_dt, end_dt, fmt, label = _get_report_params_range(request)
    qs = orders_between_by_batch_created(start_dt, end_dt)
    filename = f"cashflow_plan_{label}.{fmt}"
    content = b""
    ctype = "text/csv" if fmt == "csv" else "application/pdf"
    return _download_response(content, filename, ctype)