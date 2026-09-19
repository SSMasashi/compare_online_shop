"""
Amazon vs 楽天 最安振り分け計算
ブラウザ実行版 / Streamlit

実行方法:
    pip install streamlit

Streamlit Cloud の Secrets に以下を設定:

    RAKUTEN_APP_ID = "あなたの楽天アプリケーションID"
    RAKUTEN_ACCESS_KEY = "あなたの楽天アクセスキー"
    RAKUTEN_REFERER = "楽天APIに登録したURL"

ローカルで実行:
    streamlit run buy.py

※楽天APIの認証情報はアプリ画面には表示されません。
※保存JSONにも楽天APIの認証情報は保存されません。
"""


import os
import re
import urllib.parse
import urllib.request
import urllib.error

from itertools import product

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# ===========================================================================
# Streamlit設定
# ===========================================================================

st.set_page_config(
    page_title="Amazon×楽天 最安振り分け計算",
    page_icon="🛒",
    layout="wide",
)


# ===========================================================================
# 楽天API設定
# ===========================================================================
#
# 優先順位:
#   1. Streamlit Cloud Secrets
#   2. 環境変数
#
# Secretsの内容は画面に表示しない。
#
# ---------------------------------------------------------------------------

def get_secret_or_env(name: str, default: str = "") -> str:
    """
    Streamlit Secretsから値を取得する。

    Streamlit Cloudでは st.secrets を使用する。
    ローカル環境などでSecretsがない場合は環境変数を使用する。

    取得した値を画面には表示しない。
    """

    try:

        value = st.secrets.get(
            name,
            None
        )

        if value is not None:

            return str(
                value
            ).strip()

    except Exception:

        pass

    return os.environ.get(
        name,
        default
    ).strip()


RAKUTEN_APP_ID = get_secret_or_env(
    "RAKUTEN_APP_ID"
)

RAKUTEN_ACCESS_KEY = get_secret_or_env(
    "RAKUTEN_ACCESS_KEY"
)

RAKUTEN_REFERER = get_secret_or_env(
    "RAKUTEN_REFERER"
)


RAKUTEN_API_URL = (
    "https://openapi.rakuten.co.jp/ichibams/api/"
    "IchibaItem/Search/20260701"
)


# ===========================================================================
# Google Sheets設定
# ===========================================================================

GOOGLE_SHEET_ID = get_secret_or_env(
    "GOOGLE_SHEET_ID"
)

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PRODUCT_HEADERS = [
    "name", "ap", "apt", "baby", "rp", "rpt", "rurl"
]

SETTING_DEFAULTS = {
    "max_shops": 10,
    "min_shop_price": 1000,
    "bonus_cap": 7000,
    "spu_multiplier": 0,
}

def get_google_credentials():
    try:
        info = dict(st.secrets["gcp_service_account"])
    except Exception as e:
        raise ValueError(
            "gcp_service_accountの設定が正しくありません。"
        ) from e

    required = [
        "type", "project_id", "private_key_id", "private_key",
        "client_email", "client_id", "token_uri",
    ]
    missing = [key for key in required if not str(info.get(key, "")).strip()]
    if missing:
        raise ValueError(
            "gcp_service_accountに必要な項目がありません："
            + ", ".join(missing)
        )

    private_key = str(info["private_key"])
    if "BEGIN PRIVATE KEY" not in private_key:
        raise ValueError(
            "gcp_service_accountのprivate_keyが正しくありません。"
            "Google Cloudから発行した秘密鍵を設定してください。"
        )

    return Credentials.from_service_account_info(
        info,
        scopes=GOOGLE_SCOPES,
    )

def get_google_spreadsheet():
    if not GOOGLE_SHEET_ID:
        raise ValueError("GOOGLE_SHEET_IDが設定されていません。")

    try:
        credentials = get_google_credentials()
        client = gspread.authorize(credentials)
        return client.open_by_key(GOOGLE_SHEET_ID)
    except Exception as e:
        raise ConnectionError(
            "Googleスプレッドシートを開けませんでした。\n\n"
            "以下を確認してください。\n"
            "・GOOGLE_SHEET_IDが正しい\n"
            "・サービスアカウントにスプレッドシートを共有している\n"
            "・サービスアカウントに編集権限がある"
        ) from e

def get_or_create_worksheet(spreadsheet, title, rows=100, cols=10):
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(
            title=title,
            rows=rows,
            cols=cols,
        )

def normalize_bool(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "1", "yes", "on"}

def create_google_sheet_data():
    products = []
    for item in st.session_state.products:
        products.append([
            item.get("name", ""),
            int(item.get("ap", 0)),
            int(item.get("apt", 1)),
            bool(item.get("baby", False)),
            int(item.get("rp", 0)),
            int(item.get("rpt", 0)),
            item.get("rurl", ""),
        ])

    settings = [
        ["setting", "value"],
        ["max_shops", int(st.session_state.max_shops)],
        ["min_shop_price", int(st.session_state.min_shop_price)],
        ["bonus_cap", int(st.session_state.bonus_cap)],
        ["spu_multiplier", int(st.session_state.spu_multiplier)],
    ]
    return products, settings

def save_data_to_google_sheets():
    spreadsheet = get_google_spreadsheet()
    products_ws = get_or_create_worksheet(
        spreadsheet, "products", rows=100, cols=len(PRODUCT_HEADERS)
    )
    settings_ws = get_or_create_worksheet(
        spreadsheet, "settings", rows=20, cols=2
    )

    products, settings = create_google_sheet_data()

    product_values = [PRODUCT_HEADERS] + products
    products_ws.clear()
    products_ws.resize(
        rows=max(100, len(product_values) + 5),
        cols=len(PRODUCT_HEADERS),
    )
    products_ws.update(
        "A1",
        product_values,
        value_input_option="USER_ENTERED",
    )

    settings_ws.clear()
    settings_ws.resize(rows=20, cols=2)
    settings_ws.update(
        "A1",
        settings,
        value_input_option="USER_ENTERED",
    )

def load_data_from_google_sheets():
    spreadsheet = get_google_spreadsheet()

    products_ws = get_or_create_worksheet(
        spreadsheet, "products", rows=100, cols=len(PRODUCT_HEADERS)
    )
    settings_ws = get_or_create_worksheet(
        spreadsheet, "settings", rows=20, cols=2
    )

    # 設定は「setting,value」の見出しに依存せず読み込む
    setting_map = dict(SETTING_DEFAULTS)
    for row in settings_ws.get_all_values():
        if len(row) >= 2 and row[0] in SETTING_DEFAULTS:
            try:
                setting_map[row[0]] = int(float(row[1]))
            except (TypeError, ValueError):
                pass

    st.session_state.max_shops = setting_map["max_shops"]
    st.session_state.min_shop_price = setting_map["min_shop_price"]
    st.session_state.bonus_cap = setting_map["bonus_cap"]
    st.session_state.spu_multiplier = setting_map["spu_multiplier"]

    values = products_ws.get_all_values()
    loaded_products = []

    if values:
        header = [str(x).strip() for x in values[0]]
        try:
            indexes = {name: header.index(name) for name in PRODUCT_HEADERS}
        except ValueError:
            indexes = None

        if indexes is not None:
            for row in values[1:]:
                if not any(str(x).strip() for x in row):
                    continue
                def cell(name, default=""):
                    idx = indexes[name]
                    return row[idx] if idx < len(row) else default
                loaded_products.append({
                    "name": cell("name", ""),
                    "ap": int(float(cell("ap", 0) or 0)),
                    "apt": int(float(cell("apt", 1) or 1)),
                    "baby": normalize_bool(cell("baby", False)),
                    "rp": int(float(cell("rp", 0) or 0)),
                    "rpt": int(float(cell("rpt", 0) or 0)),
                    "rurl": cell("rurl", ""),
                })

    if loaded_products:
        st.session_state.products = loaded_products
    else:
        st.session_state.products = [{
            "name": "",
            "ap": 0,
            "apt": 1,
            "baby": False,
            "rp": 0,
            "rpt": 0,
            "rurl": "",
        }]

    st.session_state.widget_version = (
        int(st.session_state.get("widget_version", 0)) + 1
    )


# ===========================================================================
# 楽天API
# ===========================================================================

def extract_rakuten_item_code(url: str) -> str:
    """
    楽天商品URLからitemCodeを作成する。
    """

    path = urllib.parse.urlparse(
        url
    ).path.strip("/")

    parts = [
        p
        for p in path.split("/")
        if p
    ]

    if len(parts) < 2:

        raise ValueError(
            "楽天の商品URLとして認識できません"
        )

    return f"{parts[0]}:{parts[1]}"


def extract_shop_and_slug(url: str):
    """
    楽天商品URLからショップコードと商品コード部分を取得する。
    """

    path = urllib.parse.urlparse(
        url
    ).path.strip("/")

    parts = [
        p
        for p in path.split("/")
        if p
    ]

    if len(parts) < 2:

        raise ValueError(
            "楽天の商品URLとして認識できません"
        )

    return parts[0], parts[1]


def _call_rakuten_api(
    params,
    app_id,
    access_key,
    referer,
):
    """
    楽天APIを呼び出す。
    """

    params = {
        **params,
        "format": "json",
        "applicationId": app_id,
        "accessKey": access_key,
    }

    req = urllib.request.Request(
        f"{RAKUTEN_API_URL}?"
        f"{urllib.parse.urlencode(params)}",
        headers={
            "Origin": referer
        },
    )

    try:

        with urllib.request.urlopen(
            req,
            timeout=10
        ) as res:

            return json.loads(
                res.read().decode(
                    "utf-8"
                )
            )

    except urllib.error.HTTPError as e:

        body = e.read().decode(
            "utf-8",
            errors="ignore"
        )

        try:

            err = json.loads(
                body
            )

            errs = err.get(
                "errors",
                err
            )

            msg = (
                errs.get("errorMessage")
                or errs.get("error_description")
                or errs.get("error")
                or body
            )

        except json.JSONDecodeError:

            msg = body

        raise RuntimeError(
            str(msg)
        ) from None


def fetch_rakuten_price_and_point(
    url: str,
):
    """
    楽天の商品URLから価格とポイント還元率を取得する。

    楽天APIの認証情報はStreamlit Secretsから取得する。
    アプリ画面からは入力・編集できない。

    戻り値:
        price
        point_rate

    ※ここでは楽天APIから取得した生の還元率を返す。
    """

    app_id = (
        RAKUTEN_APP_ID or ""
    ).strip()

    access_key = (
        RAKUTEN_ACCESS_KEY or ""
    ).strip()

    referer = (
        RAKUTEN_REFERER or ""
    ).strip()

    if not app_id or not access_key:

        raise RuntimeError(
            "楽天APIの認証情報がStreamlit Secretsに設定されていません。"
        )

    if not referer:

        raise RuntimeError(
            "楽天APIのRAKUTEN_REFERERが"
            "Streamlit Secretsに設定されていません。"
        )

    shop_code, slug = (
        extract_shop_and_slug(
            url
        )
    )

    item_code = (
        f"{shop_code}:{slug}"
    )

    # -----------------------------------------------------------------------
    # itemCodeで直接取得
    # -----------------------------------------------------------------------

    try:

        data = _call_rakuten_api(
            {
                "itemCode": item_code
            },
            app_id,
            access_key,
            referer,
        )

        items = data.get(
            "Items",
            []
        )

        if items:

            item = items[0]["Item"]

            return (
                float(
                    item["itemPrice"]
                ),
                int(
                    float(
                        item.get(
                            "pointRate",
                            1
                        )
                    )
                ),
            )

    except RuntimeError:

        pass

    # -----------------------------------------------------------------------
    # フォールバック
    # -----------------------------------------------------------------------

    keyword_match = re.sub(
        r"[^0-9A-Za-z]+",
        " ",
        slug,
    ).split()

    keyword = (
        keyword_match[0]
        if keyword_match
        else ""
    )

    if len(keyword) < 2:

        raise ValueError(
            f"商品が見つかりませんでした"
            f"（itemCode: {item_code}）"
        )

    try:

        data = _call_rakuten_api(
            {
                "shopCode": shop_code,
                "keyword": keyword,
                "hits": 1,
            },
            app_id,
            access_key,
            referer,
        )

    except RuntimeError as e:

        raise RuntimeError(
            f"楽天APIエラー（{item_code}）: {e}"
        ) from None

    items = data.get(
        "Items",
        []
    )

    if not items:

        raise ValueError(
            f"商品が見つかりませんでした"
            f"（itemCode: {item_code} / "
            f"keyword: {keyword}）"
        )

    item = items[0]["Item"]

    return (
        float(
            item["itemPrice"]
        ),
        int(
            float(
                item.get(
                    "pointRate",
                    1
                )
            )
        ),
    )


# ===========================================================================
# 金額計算
# ===========================================================================

def tax_excluded_price(
    price,
    tax_rate=0.10,
):
    """
    税込価格から税抜価格を計算する。
    """

    return price / (
        1 + tax_rate
    )


def amazon_cost(item):
    """
    Amazonの実質価格を計算。

    らくベビ割対象の場合は10%OFF。

    Amazon還元ポイントは、
    10%OFF後の価格を基準に計算。
    """

    price = (
        item["ap"] * 0.9
        if item["baby"]
        else item["ap"]
    )

    points = (
        price
        * (
            item["apt"]
            / 100
        )
    )

    return price, points


# ===========================================================================
# 楽天買いまわり
# ===========================================================================

def calculate_rakuten_shop_count(
    items,
    choices,
    min_shop_price,
):
    """
    楽天で購入する商品のうち、
    買いまわり対象金額以上の商品数を返す。
    """

    eligible = 0

    for item, choice in zip(
        items,
        choices,
    ):

        if choice != "R":
            continue

        if item["rp"] >= min_shop_price:

            eligible += 1

    return eligible


def calculate_rakuten_bonus(
    rakuten_tax_included_total,
    eligible_shops,
    max_shops,
    bonus_cap,
):
    """
    楽天買いまわり特典ポイントを計算。
    """

    shop_count = min(
        eligible_shops,
        max_shops,
    )

    bonus_multiplier = max(
        shop_count - 1,
        0,
    )

    rakuten_tax_excluded_total = (
        tax_excluded_price(
            rakuten_tax_included_total
        )
    )

    bonus = (
        rakuten_tax_excluded_total
        * bonus_multiplier
        / 100
    )

    bonus = min(
        bonus,
        bonus_cap,
    )

    return {
        "shop_count": shop_count,

        "bonus_multiplier":
            bonus_multiplier,

        "tax_excluded_total":
            rakuten_tax_excluded_total,

        "bonus":
            bonus,
    }


# ===========================================================================
# 総合計算
# ===========================================================================

def evaluate(
    items,
    choices,
    max_shops,
    min_shop_price,
    bonus_cap,
    spu_multiplier,
):
    """
    指定された購入先の組み合わせについて
    実質負担額を計算する。
    """

    amazon_paid = 0.0
    amazon_points = 0.0

    rakuten_paid = 0.0
    rakuten_base_points = 0.0

    for item, choice in zip(
        items,
        choices,
    ):

        # -------------------------------------------------------------------
        # Amazon
        # -------------------------------------------------------------------

        if choice == "A":

            paid, points = amazon_cost(
                item
            )

            amazon_paid += paid
            amazon_points += points

        # -------------------------------------------------------------------
        # 楽天
        # -------------------------------------------------------------------

        else:

            rakuten_paid += item["rp"]

            tax_excluded = (
                tax_excluded_price(
                    item["rp"]
                )
            )

            effective_rakuten_rate = (
                item["rpt"]
                + spu_multiplier
            )

            rakuten_base_points += (
                tax_excluded
                * (
                    effective_rakuten_rate
                    / 100
                )
            )

    # -----------------------------------------------------------------------
    # 楽天買いまわり
    # -----------------------------------------------------------------------

    eligible_shops = (
        calculate_rakuten_shop_count(
            items,
            choices,
            min_shop_price,
        )
    )

    rakuten_info = (
        calculate_rakuten_bonus(
            rakuten_tax_included_total=rakuten_paid,
            eligible_shops=eligible_shops,
            max_shops=max_shops,
            bonus_cap=bonus_cap,
        )
    )

    # -----------------------------------------------------------------------
    # 合計
    # -----------------------------------------------------------------------

    total_paid = (
        amazon_paid
        + rakuten_paid
    )

    total_points = (
        amazon_points
        + rakuten_base_points
        + rakuten_info["bonus"]
    )

    net = (
        total_paid
        - total_points
    )

    return {
        "net":
            net,

        "shop_count":
            rakuten_info["shop_count"],

        "bonus_multiplier":
            rakuten_info["bonus_multiplier"],

        "bonus":
            rakuten_info["bonus"],

        "total_points":
            total_points,

        "amazon_points":
            amazon_points,

        "rakuten_base_points":
            rakuten_base_points,

        "amazon_paid":
            amazon_paid,

        "rakuten_paid":
            rakuten_paid,

        "rakuten_tax_excluded_total":
            rakuten_info[
                "tax_excluded_total"
            ],

        "spu_multiplier":
            spu_multiplier,
    }


# ===========================================================================
# 最適解探索
# ===========================================================================

def find_best(
    items,
    max_shops,
    min_shop_price,
    bonus_cap,
    spu_multiplier,
):
    """
    Amazon / 楽天の全組み合わせを調べ、
    実質負担額が最も安い組み合わせを探す。
    """

    best = None
    best_choices = None

    for choices in product(
        "AR",
        repeat=len(items),
    ):

        result = evaluate(
            items,
            choices,
            max_shops,
            min_shop_price,
            bonus_cap,
            spu_multiplier,
        )

        if (
            best is None
            or result["net"]
            < best["net"]
        ):

            best = result
            best_choices = choices

    return (
        best_choices,
        best,
    )


# ===========================================================================
# 表示用関数
# ===========================================================================

def yen(v):
    """
    円表示。
    """

    return f"{round(v):,}円"


def point(v):
    """
    ポイント表示。
    """

    return f"{round(v):,}pt"


def get_product_multiplier(
    item,
    choice,
    spu_multiplier,
    buyaround_multiplier,
    min_shop_price,
):
    """
    商品ごとの還元倍率を返す。

    Amazon:
        商品のAmazon還元率

    楽天:
        商品ポイント
        + SPU
        + 買いまわり
    """

    if choice == "A":

        total = int(
            item["apt"]
        )

        detail = (
            f"{item['apt']}"
        )

        return total, detail

    # -----------------------------------------------------------------------
    # 楽天
    # -----------------------------------------------------------------------

    product_rate = int(
        item["rpt"]
    )

    spu_rate = int(
        spu_multiplier
    )

    is_eligible = (
        item["rp"]
        >= min_shop_price
    )

    buyaround_rate = (
        int(
            buyaround_multiplier
        )
        if is_eligible
        else 0
    )

    total = (
        product_rate
        + spu_rate
        + buyaround_rate
    )

    detail = (
        f"{total}"
        f"({product_rate}"
        f"+{spu_rate}"
        f"+{buyaround_rate})"
    )

    return total, detail


# ===========================================================================
# 商品ごとのポイント計算
# ===========================================================================

def calculate_item_results(
    items,
    choices,
    best,
    spu_multiplier,
):
    """
    商品ごとの

        支払価格
        還元ポイント
        実質負担額
        還元倍率

    を計算する。

    楽天買いまわりポイントは、
    全楽天商品の税抜価格に応じて比例配分する。

    これにより、

        商品ごとの還元ポイント合計
        =
        全体のポイント合計

    となる。
    """

    results = []

    # -----------------------------------------------------------------------
    # 楽天商品の税抜価格合計
    # -----------------------------------------------------------------------

    rakuten_tax_excluded_total = 0.0

    for item, choice in zip(
        items,
        choices,
    ):

        if choice == "R":

            rakuten_tax_excluded_total += (
                tax_excluded_price(
                    item["rp"]
                )
            )

    # -----------------------------------------------------------------------
    # 買いまわりポイント
    # -----------------------------------------------------------------------

    total_rakuten_bonus = float(
        best["bonus"]
    )

    # -----------------------------------------------------------------------
    # 商品ごとの計算
    # -----------------------------------------------------------------------

    for item, choice in zip(
        items,
        choices,
    ):

        name = (
            item["name"]
            or "(無題)"
        )

        # ===================================================================
        # Amazon
        # ===================================================================

        if choice == "A":

            paid, base_points = (
                amazon_cost(
                    item
                )
            )

            total_points = (
                base_points
            )

            net = (
                paid
                - total_points
            )

            multiplier = int(
                item["apt"]
            )

            results.append(
                {
                    "name":
                        name,

                    "store":
                        "🟧 Amazon",

                    "price":
                        paid,

                    "multiplier":
                        f"{multiplier}倍",

                    "points":
                        total_points,

                    "net":
                        net,
                }
            )

        # ===================================================================
        # 楽天
        # ===================================================================

        else:

            paid = float(
                item["rp"]
            )

            tax_excluded = (
                tax_excluded_price(
                    paid
                )
            )

            # ---------------------------------------------------------------
            # 商品ポイント + SPU
            # ---------------------------------------------------------------

            product_rate = float(
                item["rpt"]
            )

            spu_rate = float(
                spu_multiplier
            )

            base_points = (
                tax_excluded
                * (
                    product_rate
                    + spu_rate
                )
                / 100
            )

            # ---------------------------------------------------------------
            # 買いまわりポイント
            #
            # 現在の全体計算ロジックでは、
            # 楽天購入商品の税抜合計を基準にしているため、
            # 各楽天商品の税抜価格に比例して配分する。
            # ---------------------------------------------------------------

            buyaround_points = 0.0

            if (
                rakuten_tax_excluded_total > 0
                and total_rakuten_bonus > 0
            ):

                buyaround_points = (
                    total_rakuten_bonus
                    * tax_excluded
                    / rakuten_tax_excluded_total
                )

            total_points = (
                base_points
                + buyaround_points
            )

            net = (
                paid
                - total_points
            )

            eligible = (
                paid
                >= st.session_state.min_shop_price
            )

            buyaround_rate = (
                best["bonus_multiplier"]
                if eligible
                else 0
            )

            total_multiplier = (
                product_rate
                + spu_rate
                + buyaround_rate
            )

            detail = (
                f"{round(total_multiplier)}倍"
                f"("
                f"{round(product_rate)}"
                f"+"
                f"{round(spu_rate)}"
                f"+"
                f"{round(buyaround_rate)}"
                f")"
            )

            results.append(
                {
                    "name":
                        name,

                    "store":
                        "🟥 楽天",

                    "price":
                        paid,

                    "multiplier":
                        detail,

                    "points":
                        total_points,

                    "net":
                        net,
                }
            )

    return results


# ===========================================================================
# 保存・読み込み
# ===========================================================================

# Google Sheetsの読み込みはウィジェット生成前に実行する。
# これにより、読み込んだ値がそのまま画面の入力欄に反映される。
if st.session_state.pop("load_requested", False):
    try:
        load_data_from_google_sheets()
        st.session_state["google_sheet_message"] = (
            "success", "Google Sheetsから設定を読み込みました。"
        )
    except Exception as e:
        st.session_state["google_sheet_message"] = (
            "error", f"Google Sheetsからの読み込みに失敗しました：{e}"
        )


# ===========================================================================
# 初期値
# ===========================================================================

if "products" not in st.session_state:

    st.session_state.products = [

        {
            "name": "商品A",
            "ap": 3000,
            "apt": 1,
            "baby": False,
            "rp": 3200,
            "rpt": 5,
            "rurl": "",
        },

        {
            "name": "商品B",
            "ap": 5000,
            "apt": 1,
            "baby": True,
            "rp": 4800,
            "rpt": 8,
            "rurl": "",
        },
    ]


if "widget_version" not in st.session_state:

    st.session_state.widget_version = 0


if "max_shops" not in st.session_state:

    st.session_state.max_shops = 10


if "min_shop_price" not in st.session_state:

    st.session_state.min_shop_price = 1000


if "bonus_cap" not in st.session_state:

    st.session_state.bonus_cap = 7000


if "spu_multiplier" not in st.session_state:

    st.session_state.spu_multiplier = 0


# ===========================================================================
# タイトル
# ===========================================================================

st.title(
    "🛒 Amazon × 楽天 最安振り分け計算"
)

st.caption(
    "各商品をAmazonと楽天市場のどちらで購入すると、"
    "ポイントを含めた実質負担額が安くなるかを計算します。"
)


# ===========================================================================
# 設定
# ===========================================================================
#
# 楽天APIの認証情報はここには表示しない。
#
# ===========================================================================

with st.expander(
    "⚙️ 設定（楽天買いまわり）",
    expanded=False,
):

    v = st.session_state.widget_version

    c1, c2, c3, c4 = st.columns(
        4
    )

    with c1:

        st.session_state.max_shops = (
            st.number_input(
                "買いまわり最大ショップ数",
                min_value=1,
                max_value=10,
                value=int(
                    st.session_state.max_shops
                ),
                step=1,
                format="%d",
                key=f"max_shops_widget_{v}",
            )
        )

    with c2:

        st.session_state.min_shop_price = (
            st.number_input(
                "買いまわり対象最低金額（税込）",
                min_value=0,
                value=int(
                    st.session_state.min_shop_price
                ),
                step=100,
                format="%d",
                key=f"min_shop_price_widget_{v}",
            )
        )

    with c3:

        st.session_state.bonus_cap = (
            st.number_input(
                "買いまわり特典ポイント上限",
                min_value=0,
                value=int(
                    st.session_state.bonus_cap
                ),
                step=100,
                format="%d",
                key=f"bonus_cap_widget_{v}",
            )
        )

    with c4:

        st.session_state.spu_multiplier = (
            st.number_input(
                "楽天SPUポイント倍率",
                min_value=0,
                max_value=100,
                value=int(
                    st.session_state.spu_multiplier
                ),
                step=1,
                format="%d",
                key=f"spu_multiplier_widget_{v}",
                help=(
                    "楽天の商品ごとの還元率に加算します。"
                    "例：3と設定すると+3倍。"
                ),
            )
        )

    st.caption(
        "※SPUは計算時のみ楽天還元率に加算されます。"
        "商品リストの楽天還元%には表示されません。"
    )

    st.caption(
        "※楽天APIから取得した還元率は1倍引いて表示します。"
    )

    st.caption(
        "※楽天APIの認証情報はアプリ画面には表示されません。"
        "Streamlit Cloud Secretsから取得します。"
    )


# ===========================================================================
# 商品リスト
# ===========================================================================

st.subheader(
    "商品リスト"
)

items = st.session_state.products


for i, item in enumerate(
    items
):

    v = st.session_state.widget_version

    # =======================================================================
    # 商品名・削除
    # =======================================================================

    name_col, delete_col = st.columns(
        [9, 1]
    )

    with name_col:

        current_name = item["name"]

        item["name"] = st.text_input(
            "",
            value=current_name,
            placeholder="商品名",
            key=f"name_{i}_{v}",
            label_visibility="collapsed",
        )

    with delete_col:

        if st.button(
            "🗑️",
            key=f"delete_{i}_{v}",
        ):

            st.session_state.products.pop(
                i
            )

            st.session_state.widget_version += 1

            st.rerun()

    # =======================================================================
    # Amazon
    # =======================================================================

    (
        amazon_label,
        amazon_price_label,
        amazon_price,
        amazon_return_label,
        amazon_return,
        baby_col,
        amazon_empty,
    ) = st.columns(
        [
            1.4,
            0.7,
            1.5,
            0.7,
            1.5,
            4.2,
            1.6,
        ]
    )

    with amazon_label:

        st.markdown(
            "**🟧 Amazon**"
        )

    with amazon_price_label:

        st.markdown(
            "価格"
        )

    with amazon_price:

        item["ap"] = st.number_input(
            "",
            min_value=0,
            value=int(
                item.get(
                    "ap",
                    0
                )
            ),
            step=100,
            key=f"ap_{i}_{v}",
            label_visibility="collapsed",
        )

    with amazon_return_label:

        st.markdown(
            "還元%"
        )

    with amazon_return:

        item["apt"] = st.number_input(
            "",
            min_value=0,
            max_value=100,
            value=int(
                item.get(
                    "apt",
                    1
                )
            ),
            step=1,
            format="%d",
            key=f"apt_{i}_{v}",
            label_visibility="collapsed",
        )

    with baby_col:

        item["baby"] = st.checkbox(
            "らくベビ割（10%OFF）",
            value=item.get(
                "baby",
                False
            ),
            key=f"baby_{i}_{v}",
        )

    # =======================================================================
    # 楽天
    # =======================================================================

    (
        rakuten_label,
        rakuten_price_label,
        rakuten_price,
        rakuten_return_label,
        rakuten_return,
        rakuten_url_input,
        rakuten_url_button,
    ) = st.columns(
        [
            1.4,
            0.7,
            1.5,
            0.7,
            1.5,
            4.2,
            1.6,
        ]
    )

    with rakuten_label:

        st.markdown(
            "**🟥 楽天**"
        )

    with rakuten_price_label:

        st.markdown(
            "価格"
        )

    with rakuten_price:

        item["rp"] = st.number_input(
            "",
            min_value=0,
            value=int(
                item.get(
                    "rp",
                    0
                )
            ),
            step=100,
            key=f"rp_{i}_{v}",
            label_visibility="collapsed",
        )

    with rakuten_return_label:

        st.markdown(
            "還元%"
        )

    with rakuten_return:

        item["rpt"] = st.number_input(
            "",
            min_value=0,
            max_value=100,
            value=int(
                item.get(
                    "rpt",
                    0
                )
            ),
            step=1,
            format="%d",
            key=f"rpt_{i}_{v}",
            label_visibility="collapsed",
        )

    with rakuten_url_input:

        item["rurl"] = st.text_input(
            "",
            value=item.get(
                "rurl",
                ""
            ),
            placeholder="URL",
            key=f"rurl_{i}_{v}",
            label_visibility="collapsed",
        )

    with rakuten_url_button:

        if st.button(
            "取得",
            key=f"get_rakuten_{i}_{v}",
            use_container_width=True,
        ):

            url = st.session_state.get(
                f"rurl_{i}_{v}",
                "",
            ).strip()

            if not url:

                st.session_state[
                    f"rakuten_message_{i}"
                ] = (
                    "楽天商品URLを入力してください"
                )

            else:

                try:

                    price, point_rate = (
                        fetch_rakuten_price_and_point(
                            url
                        )
                    )

                    # -------------------------------------------------------
                    # 楽天価格
                    # -------------------------------------------------------

                    display_price = int(
                        round(price)
                    )

                    # -------------------------------------------------------
                    # 楽天還元率
                    #
                    # API取得値から1を引く
                    #
                    # API 4 → 表示 3
                    # API 3 → 表示 2
                    # API 1 → 表示 0
                    # -------------------------------------------------------

                    display_point_rate = max(
                        int(
                            point_rate
                        ) - 1,
                        0,
                    )

                    # -------------------------------------------------------
                    # 商品データを更新
                    # -------------------------------------------------------

                    item["rp"] = (
                        display_price
                    )

                    item["rpt"] = (
                        display_point_rate
                    )

                    item["rurl"] = url

                    st.session_state[
                        f"rakuten_message_{i}"
                    ] = (
                        "楽天の商品情報を取得しました"
                    )

                    st.session_state.widget_version += 1

                    st.rerun()

                except Exception as e:

                    st.session_state[
                        f"rakuten_message_{i}"
                    ] = (
                        f"取得エラー：{e}"
                    )

    # -----------------------------------------------------------------------
    # URL取得メッセージ
    # -----------------------------------------------------------------------

    message = st.session_state.get(
        f"rakuten_message_{i}"
    )

    if message:

        if message.startswith(
            "取得エラー"
        ):

            st.error(
                message
            )

        elif message.startswith(
            "楽天商品URLを入力"
        ):

            st.warning(
                message
            )

        else:

            st.success(
                message
            )

    st.divider()


# ===========================================================================
# 商品追加
# ===========================================================================

col_a, col_b = st.columns(
    2
)

with col_a:

    if st.button(
        "＋ 商品を追加"
    ):

        if len(
            st.session_state.products
        ) < 15:

            st.session_state.products.append(
                {
                    "name": "",

                    "ap": 0,

                    # ★新規商品のAmazon還元率は1%
                    "apt": 1,

                    "baby": False,

                    "rp": 0,

                    "rpt": 0,

                    "rurl": "",
                }
            )

            st.session_state.widget_version += 1

            st.rerun()

        else:

            st.warning(
                "商品は最大15個までです。"
            )


# ===========================================================================
# 保存・読み込み
# ===========================================================================

st.subheader(
    "💾 保存・読み込み"
)

save_col, load_col, info_col = st.columns(
    [2, 2, 4]
)

with save_col:
    if st.button(
        "💾 現在の設定を保存",
        use_container_width=True,
    ):
        try:
            save_data_to_google_sheets()
            st.session_state["google_sheet_message"] = (
                "success", "Google Sheetsへ保存しました。"
            )
        except Exception as e:
            st.session_state["google_sheet_message"] = (
                "error", f"Google Sheetsへの保存に失敗しました：{e}"
            )

with load_col:
    if st.button(
        "📂 保存データを読み込む",
        use_container_width=True,
    ):
        st.session_state["load_requested"] = True
        st.rerun()

with info_col:
    st.caption(
        "保存先：Google Sheets"
    )
    st.caption(
        "楽天APIの認証情報はGoogle Sheetsには保存されません。"
    )

# Google Sheets関連メッセージは、保存・読み込みボタンの直下に表示する。
google_sheet_message = st.session_state.pop(
    "google_sheet_message",
    None,
)

if google_sheet_message:
    message_type, message_text = google_sheet_message
    if message_type == "success":
        st.success(message_text)
    elif message_type == "warning":
        st.warning(message_text)
    else:
        st.error(message_text)


# ===========================================================================
# 計算
# ===========================================================================

if st.button(
    "🧮 計算する",
    type="primary",
    use_container_width=True,
):

    items = st.session_state.products

    # -----------------------------------------------------------------------
    # すべてAmazon
    # -----------------------------------------------------------------------

    all_a = evaluate(
        items,
        ["A"] * len(items),
        st.session_state.max_shops,
        st.session_state.min_shop_price,
        st.session_state.bonus_cap,
        st.session_state.spu_multiplier,
    )

    # -----------------------------------------------------------------------
    # すべて楽天
    # -----------------------------------------------------------------------

    all_r = evaluate(
        items,
        ["R"] * len(items),
        st.session_state.max_shops,
        st.session_state.min_shop_price,
        st.session_state.bonus_cap,
        st.session_state.spu_multiplier,
    )

    # -----------------------------------------------------------------------
    # 最適な振り分け
    # -----------------------------------------------------------------------

    best_choices, best = find_best(
        items,
        st.session_state.max_shops,
        st.session_state.min_shop_price,
        st.session_state.bonus_cap,
        st.session_state.spu_multiplier,
    )

    # -----------------------------------------------------------------------
    # 商品ごとの結果を計算
    # -----------------------------------------------------------------------

    item_results = calculate_item_results(
        items,
        best_choices,
        best,
        st.session_state.spu_multiplier,
    )

    # -----------------------------------------------------------------------
    # 結果
    # -----------------------------------------------------------------------

    st.subheader(
        "🧮 計算結果"
    )

    result_col1, result_col2, result_col3 = st.columns(
        3
    )

    with result_col1:

        st.metric(
            "⭐ 最適な振り分け",
            yen(best["net"]),
        )

    with result_col2:

        st.metric(
            "🟧 すべてAmazon",
            yen(all_a["net"]),
        )

    with result_col3:

        st.metric(
            "🟥 すべて楽天",
            yen(all_r["net"]),
        )

    # =========================================================================
    # 商品ごとの購入先
    # =========================================================================

    st.markdown(
        "### 📦 商品ごとの購入先"
    )

    # -------------------------------------------------------------------------
    # ヘッダー
    #
    # 左から:
    #   商品名
    #   購入先
    #   価格
    #   還元倍率
    #   還元ポイント
    #   実質負担額
    # -------------------------------------------------------------------------

    (
        h_name,
        h_store,
        h_price,
        h_multiplier,
        h_points,
        h_net,
    ) = st.columns(
        [
            2.4,
            1.4,
            1.5,
            2.1,
            1.6,
            1.7,
        ]
    )

    with h_name:

        st.markdown(
            "**商品名**"
        )

    with h_store:

        st.markdown(
            "**購入先**"
        )

    with h_price:

        st.markdown(
            "**価格**"
        )

    with h_multiplier:

        st.markdown(
            "**還元倍率**"
        )

    with h_points:

        st.markdown(
            "**還元ポイント**"
        )

    with h_net:

        st.markdown(
            "**実質負担額**"
        )

    # -------------------------------------------------------------------------
    # 商品ごとの結果
    # -------------------------------------------------------------------------

    for result in item_results:

        (
            row_name,
            row_store,
            row_price,
            row_multiplier,
            row_points,
            row_net,
        ) = st.columns(
            [
                2.4,
                1.4,
                1.5,
                2.1,
                1.6,
                1.7,
            ]
        )

        with row_name:

            st.write(
                result["name"]
            )

        with row_store:

            st.write(
                result["store"]
            )

        with row_price:

            st.write(
                yen(
                    result["price"]
                )
            )

        with row_multiplier:

            st.write(
                result["multiplier"]
            )

        with row_points:

            st.write(
                point(
                    result["points"]
                )
            )

        with row_net:

            st.write(
                yen(
                    result["net"]
                )
            )

    # =========================================================================
    # 合計
    # =========================================================================

    st.divider()

    (
        total_name,
        total_store,
        total_price,
        total_multiplier,
        total_points_col,
        total_net,
    ) = st.columns(
        [
            2.4,
            1.4,
            1.5,
            2.1,
            1.6,
            1.7,
        ]
    )

    with total_name:

        st.markdown(
            "**合計**"
        )

    with total_store:

        st.markdown(
            "**購入額**"
        )

    with total_price:

        st.markdown(
            "**"
            f"{yen(best['amazon_paid'] + best['rakuten_paid'])}"
            "**"
        )

    with total_multiplier:

        st.markdown(
            "**―**"
        )

    with total_points_col:

        st.markdown(
            f"**{point(best['total_points'])}**"
        )

    with total_net:

        st.markdown(
            f"**{yen(best['net'])}**"
        )

    # =========================================================================
    # 詳細計算
    # =========================================================================

    with st.expander(
        "📊 詳細な計算内訳"
    ):

        # ---------------------------------------------------------------------
        # Amazon
        # ---------------------------------------------------------------------

        st.markdown(
            "#### 🟧 Amazon"
        )

        st.write(
            f"支払額："
            f"{yen(best['amazon_paid'])}"
        )

        st.write(
            f"Amazonポイント："
            f"{point(best['amazon_points'])}"
        )

        st.divider()

        # ---------------------------------------------------------------------
        # 楽天
        # ---------------------------------------------------------------------

        st.markdown(
            "#### 🟥 楽天市場"
        )

        st.write(
            f"税込支払額："
            f"{yen(best['rakuten_paid'])}"
        )

        st.write(
            f"税抜購入額："
            f"{yen(best['rakuten_tax_excluded_total'])}"
        )

        st.write(
            f"楽天通常ポイント："
            f"{point(best['rakuten_base_points'])}"
        )

        st.write(
            f"楽天SPU："
            f"+{best['spu_multiplier']}倍"
        )

        st.write(
            f"買いまわり対象ショップ数："
            f"{best['shop_count']}"
        )

        st.write(
            f"買いまわり特典："
            f"+{best['bonus_multiplier']}倍"
        )

        st.write(
            f"買いまわりポイント："
            f"{point(best['bonus'])}"
        )

        st.divider()

        # ---------------------------------------------------------------------
        # 最終結果
        # ---------------------------------------------------------------------

        st.markdown(
            "#### 💰 最終結果"
        )

        st.write(
            f"支払額合計："
            f"{yen(best['amazon_paid'] + best['rakuten_paid'])}"
        )

        st.write(
            f"ポイント合計："
            f"{point(best['total_points'])}"
        )

        st.write(
            f"**実質負担額："
            f"{yen(best['net'])}**"
        )
