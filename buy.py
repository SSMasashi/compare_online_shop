import itertools
import re
from urllib.parse import urlparse, unquote

import gspread
import requests
import streamlit as st
from google.oauth2.service_account import Credentials


# ============================================================
# Streamlit 基本設定
# ============================================================

st.set_page_config(
    page_title="Amazon vs 楽天 最安振り分け計算",
    page_icon="🛒",
    layout="wide",
)


# ============================================================
# 定数
# ============================================================

RAKUTEN_API_URL = (
    "https://openapi.rakuten.co.jp/ichibams/api/"
    "IchibaItem/Search/20260701"
)

PRODUCTS_SHEET_NAME = "products"
SETTINGS_SHEET_NAME = "settings"

PRODUCT_COLUMNS = [
    "name",
    "ap",
    "apt",
    "baby",
    "rp",
    "rpt",
    "rurl",
]

DEFAULT_PRODUCT = {
    "name": "",
    "ap": 0,
    "apt": 1,
    "baby": False,
    "rp": 0,
    "rpt": 0,
    "rurl": "",
}

DEFAULT_SETTINGS = {
    "max_shops": 10,
    "min_shop_price": 1000,
    "bonus_cap": 7000,
    "spu_multiplier": 0,
}


# ============================================================
# Secrets
# ============================================================

def get_secret(name, default=None):
    """
    Streamlit Secretsから値を取得する。
    """

    try:
        value = st.secrets[name]

        if value is None:
            return default

        return value

    except Exception:
        return default


def get_rakuten_credentials():
    """
    楽天API認証情報をSecretsから取得。
    画面には表示しない。
    """

    app_id = get_secret("RAKUTEN_APP_ID")
    access_key = get_secret("RAKUTEN_ACCESS_KEY")
    referer = get_secret("RAKUTEN_REFERER")

    return app_id, access_key, referer


# ============================================================
# Google Sheets
# ============================================================

@st.cache_resource
def get_google_sheet_client():
    """
    Streamlit Secretsの

    [gcp_service_account]

    を使ってGoogle Sheetsへ接続する。
    """

    try:
        service_account_info = dict(
            st.secrets["gcp_service_account"]
        )

    except Exception as e:
        raise RuntimeError(
            "Streamlit Secretsに "
            "[gcp_service_account] "
            "が設定されていません。"
        ) from e

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=scopes,
        )
    except Exception as e:
        raise RuntimeError(
            "gcp_service_account の設定が正しくありません。"
        ) from e

    return gspread.authorize(credentials)


def get_spreadsheet():
    """
    GOOGLE_SHEET_IDで指定されたスプレッドシートを取得。
    """

    sheet_id = get_secret("GOOGLE_SHEET_ID")

    if not sheet_id:
        raise RuntimeError(
            "GOOGLE_SHEET_ID がStreamlit Secretsに設定されていません。"
        )

    client = get_google_sheet_client()

    try:
        return client.open_by_key(sheet_id)

    except Exception as e:
        raise RuntimeError(
            "Googleスプレッドシートを開けませんでした。\n\n"
            "以下を確認してください。\n"
            "・GOOGLE_SHEET_IDが正しい\n"
            "・サービスアカウントにスプレッドシートを共有している\n"
            "・サービスアカウントに編集権限がある"
        ) from e


def get_or_create_worksheet(
    spreadsheet,
    title,
    rows=100,
    cols=20,
):
    """
    指定したワークシートを取得。
    存在しなければ作成。
    """

    try:
        return spreadsheet.worksheet(title)

    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(
            title=title,
            rows=rows,
            cols=cols,
        )


# ============================================================
# Google Sheets 保存
# ============================================================

def save_data_to_google_sheets():
    """
    商品データと設定をGoogle Sheetsへ保存。

    楽天API認証情報は一切保存しない。
    """

    spreadsheet = get_spreadsheet()

    # --------------------------------------------------------
    # products
    # --------------------------------------------------------

    products_ws = get_or_create_worksheet(
        spreadsheet,
        PRODUCTS_SHEET_NAME,
        rows=100,
        cols=len(PRODUCT_COLUMNS),
    )

    product_values = [
        PRODUCT_COLUMNS
    ]

    for product in st.session_state.products:
        product_values.append([
            product.get("name", ""),
            product.get("ap", 0),
            product.get("apt", 1),
            product.get("baby", False),
            product.get("rp", 0),
            product.get("rpt", 0),
            product.get("rurl", ""),
        ])

    products_ws.clear()

    products_ws.update(
        range_name="A1",
        values=product_values,
    )

    # --------------------------------------------------------
    # settings
    # --------------------------------------------------------

    settings_ws = get_or_create_worksheet(
        spreadsheet,
        SETTINGS_SHEET_NAME,
        rows=20,
        cols=2,
    )

    settings_values = [
        ["setting", "value"],
        [
            "max_shops",
            st.session_state.max_shops,
        ],
        [
            "min_shop_price",
            st.session_state.min_shop_price,
        ],
        [
            "bonus_cap",
            st.session_state.bonus_cap,
        ],
        [
            "spu_multiplier",
            st.session_state.spu_multiplier,
        ],
    ]

    settings_ws.clear()

    settings_ws.update(
        range_name="A1",
        values=settings_values,
    )


# ============================================================
# Google Sheets 読み込み
# ============================================================

def load_data_from_google_sheets():
    """
    Google Sheetsから商品・設定を読み込む。

    読み込みに失敗した場合はFalseを返す。
    """

    try:
        spreadsheet = get_spreadsheet()

    except Exception:
        return False

    # --------------------------------------------------------
    # products
    # --------------------------------------------------------

    try:
        products_ws = spreadsheet.worksheet(
            PRODUCTS_SHEET_NAME
        )

        records = products_ws.get_all_records()

        products = []

        for row in records:
            product = {
                "name": row.get("name", ""),
                "ap": safe_float(row.get("ap", 0)),
                "apt": safe_float(row.get("apt", 1)),
                "baby": parse_bool(row.get("baby", False)),
                "rp": safe_float(row.get("rp", 0)),
                "rpt": safe_float(row.get("rpt", 0)),
                "rurl": row.get("rurl", ""),
            }

            products.append(product)

        if products:
            st.session_state.products = products

    except Exception:
        pass

    # --------------------------------------------------------
    # settings
    # --------------------------------------------------------

    try:
        settings_ws = spreadsheet.worksheet(
            SETTINGS_SHEET_NAME
        )

        records = settings_ws.get_all_records()

        setting_dict = {}

        for row in records:
            key = row.get("setting")
            value = row.get("value")

            if key:
                setting_dict[key] = value

        if "max_shops" in setting_dict:
            st.session_state.max_shops = int(
                safe_float(setting_dict["max_shops"])
            )

        if "min_shop_price" in setting_dict:
            st.session_state.min_shop_price = safe_float(
                setting_dict["min_shop_price"]
            )

        if "bonus_cap" in setting_dict:
            st.session_state.bonus_cap = safe_float(
                setting_dict["bonus_cap"]
            )

        if "spu_multiplier" in setting_dict:
            st.session_state.spu_multiplier = safe_float(
                setting_dict["spu_multiplier"]
            )

    except Exception:
        pass

    return True


# ============================================================
# Utility
# ============================================================

def safe_float(value, default=0):
    """
    数値へ安全に変換。
    """

    if value is None:
        return default

    if isinstance(value, bool):
        return 1 if value else 0

    try:
        if isinstance(value, str):
            value = (
                value
                .replace(",", "")
                .replace("円", "")
                .replace("%", "")
                .strip()
            )

        return float(value)

    except Exception:
        return default


def parse_bool(value):
    """
    Google Sheets等から取得した値をboolへ変換。
    """

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return bool(value)

    if isinstance(value, str):
        return value.strip().lower() in [
            "true",
            "1",
            "yes",
            "y",
            "on",
            "はい",
        ]

    return False


def yen(value):
    """
    円表示。
    """

    return f"{round(value):,}円"


# ============================================================
# 商品関連
# ============================================================

def normalize_product(product):
    """
    商品データを正規化。
    """

    return {
        "name": str(
            product.get("name", "")
        ),

        "ap": safe_float(
            product.get("ap", 0)
        ),

        "apt": safe_float(
            product.get("apt", 1)
        ),

        "baby": parse_bool(
            product.get("baby", False)
        ),

        "rp": safe_float(
            product.get("rp", 0)
        ),

        "rpt": safe_float(
            product.get("rpt", 0)
        ),

        "rurl": str(
            product.get("rurl", "")
        ),
    }


# ============================================================
# 楽天URL解析
# ============================================================

def extract_shop_and_slug(url):
    """
    楽天市場URLからショップ名・商品コード候補を取得。

    例:
    https://item.rakuten.co.jp/shop/itemcode/
    """

    if not url:
        return None, None

    try:
        parsed = urlparse(url)

        path_parts = [
            unquote(x)
            for x in parsed.path.split("/")
            if x
        ]

        if len(path_parts) >= 2:
            shop = path_parts[-2]
            slug = path_parts[-1]

            return shop, slug

    except Exception:
        pass

    return None, None


# ============================================================
# 楽天API
# ============================================================

def _call_rakuten_api(
    keyword=None,
    item_code=None,
):
    """
    楽天市場APIを呼び出す。
    """

    app_id, access_key, referer = get_rakuten_credentials()

    if not app_id:
        raise RuntimeError(
            "RAKUTEN_APP_ID がStreamlit Secretsにありません。"
        )

    if not access_key:
        raise RuntimeError(
            "RAKUTEN_ACCESS_KEY がStreamlit Secretsにありません。"
        )

    params = {
        "applicationId": app_id,
        "accessKey": access_key,
        "formatVersion": 2,
        "hits": 1,
    }

    if keyword:
        params["keyword"] = keyword

    if item_code:
        params["itemCode"] = item_code

    headers = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/140.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    if referer:
        headers["Referer"] = referer

    response = requests.get(
        RAKUTEN_API_URL,
        params=params,
        headers=headers,
        timeout=20,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"楽天APIエラー: HTTP {response.status_code}\n"
            f"{response.text[:500]}"
        )

    data = response.json()

    return data


def fetch_rakuten_price_and_point(
    url,
    keyword=None,
):
    """
    楽天URLまたは商品名から楽天商品情報を取得。

    戻り値:
        price
        point_rate
        item_name
        shop_name
    """

    if not url and not keyword:
        return None

    shop, slug = extract_shop_and_slug(url)

    # --------------------------------------------------------
    # 1. itemCode検索
    # --------------------------------------------------------

    if shop and slug:

        item_code = f"{shop}:{slug}"

        try:
            data = _call_rakuten_api(
                item_code=item_code
            )

            items = data.get("Items", [])

            if items:
                item = items[0]

                return parse_rakuten_item(item)

        except Exception:
            pass

    # --------------------------------------------------------
    # 2. 商品名検索
    # --------------------------------------------------------

    if keyword:

        try:
            data = _call_rakuten_api(
                keyword=keyword
            )

            items = data.get("Items", [])

            if items:
                item = items[0]

                return parse_rakuten_item(item)

        except Exception:
            pass

    return None


def parse_rakuten_item(item):
    """
    楽天APIの1商品データを解析。
    """

    # formatVersion=2ではItem配下に必要情報が入る。
    if isinstance(item, dict) and "Item" in item:
        item = item["Item"]

    price = safe_float(
        item.get("itemPrice", 0)
    )

    point_rate_raw = safe_float(
        item.get("pointRate", 0)
    )

    # 楽天APIのpointRateから通常1倍分を除外。
    point_rate = max(
        int(point_rate_raw) - 1,
        0,
    )

    item_name = item.get(
        "itemName",
        "",
    )

    item_url = item.get(
        "itemUrl",
        "",
    )

    shop_name = item.get(
        "shopName",
        "",
    )

    return {
        "price": price,
        "point_rate": point_rate,
        "item_name": item_name,
        "item_url": item_url,
        "shop_name": shop_name,
    }


# ============================================================
# 買い回り計算
# ============================================================

def calculate_buyaround_bonus(
    rakuten_products,
    max_shops,
    min_shop_price,
    bonus_cap,
):
    """
    楽天買い回りボーナスを計算。

    商品単位でショップを判定できる場合はショップ数を数える。
    ショップ情報が取れない商品については商品単位で扱う。
    """

    if not rakuten_products:
        return {
            "shop_count": 0,
            "bonus_multiplier": 0,
            "bonus_points": 0,
        }

    eligible_products = [
        p
        for p in rakuten_products
        if p["price"] >= min_shop_price
    ]

    if not eligible_products:
        return {
            "shop_count": 0,
            "bonus_multiplier": 0,
            "bonus_points": 0,
        }

    shops = set()

    for p in eligible_products:

        shop = p.get("shop")

        if shop:
            shops.add(shop)
        else:
            # ショップ情報がない場合は商品ごとに仮ID
            shops.add(
                f"product_{p['index']}"
            )

    shop_count = min(
        len(shops),
        int(max_shops),
    )

    bonus_multiplier = max(
        shop_count - 1,
        0,
    )

    if bonus_multiplier <= 0:
        return {
            "shop_count": shop_count,
            "bonus_multiplier": 0,
            "bonus_points": 0,
        }

    total_tax_excluded = sum(
        p["price"] / 1.10
        for p in eligible_products
    )

    bonus_points = (
        total_tax_excluded
        * bonus_multiplier
        / 100
    )

    bonus_points = min(
        bonus_points,
        bonus_cap,
    )

    return {
        "shop_count": shop_count,
        "bonus_multiplier": bonus_multiplier,
        "bonus_points": bonus_points,
    }


# ============================================================
# 商品ごとの計算
# ============================================================

def calculate_item_results(
    products,
    assignments,
    spu_multiplier,
    max_shops,
    min_shop_price,
    bonus_cap,
):
    """
    選択された購入先に対して商品ごとの結果を計算。
    """

    rakuten_products = []

    for index, product in enumerate(products):

        if assignments[index] != "R":
            continue

        shop = None

        if product.get("rurl"):
            shop, _ = extract_shop_and_slug(
                product["rurl"]
            )

        rakuten_products.append({
            "index": index,
            "price": product["rp"],
            "shop": shop,
        })

    buyaround = calculate_buyaround_bonus(
        rakuten_products,
        max_shops,
        min_shop_price,
        bonus_cap,
    )

    total_rakuten_tax_excluded = sum(
        products[i]["rp"] / 1.10
        for i in range(len(products))
        if assignments[i] == "R"
        and products[i]["rp"] >= min_shop_price
    )

    rows = []

    for index, product in enumerate(products):

        purchase = assignments[index]

        # ----------------------------------------------------
        # Amazon
        # ----------------------------------------------------

        if purchase == "A":

            base_price = product["ap"]

            if product["baby"]:
                price = base_price * 0.90
            else:
                price = base_price

            point_rate = product["apt"]

            points = (
                price
                * point_rate
                / 100
            )

            net = price - points

            rows.append({
                "商品名": product["name"],
                "購入先": "Amazon",
                "価格": price,
                "還元倍率": point_rate,
                "還元ポイント": points,
                "実質負担額": net,
            })

        # ----------------------------------------------------
        # Rakuten
        # ----------------------------------------------------

        else:

            price = product["rp"]

            base_point_rate = product["rpt"]

            total_point_rate = (
                1
                + base_point_rate
                + spu_multiplier
            )

            points = (
                price
                / 1.10
                * total_point_rate
                / 100
            )

            # ------------------------------------------------
            # 買い回りボーナス
            # ------------------------------------------------

            bonus_points = 0

            if (
                product["rp"] >= min_shop_price
                and total_rakuten_tax_excluded > 0
                and buyaround["bonus_points"] > 0
            ):

                product_tax_excluded = (
                    product["rp"] / 1.10
                )

                ratio = (
                    product_tax_excluded
                    / total_rakuten_tax_excluded
                )

                bonus_points = (
                    buyaround["bonus_points"]
                    * ratio
                )

            total_points = (
                points
                + bonus_points
            )

            net = price - total_points

            rows.append({
                "商品名": product["name"],
                "購入先": "楽天",
                "価格": price,
                "還元倍率": total_point_rate,
                "還元ポイント": total_points,
                "実質負担額": net,
            })

    return rows, buyaround


# ============================================================
# 合計計算
# ============================================================

def calculate_total(rows):
    """
    結果の合計。
    """

    total_price = sum(
        row["価格"]
        for row in rows
    )

    total_points = sum(
        row["還元ポイント"]
        for row in rows
    )

    total_net = sum(
        row["実質負担額"]
        for row in rows
    )

    return (
        total_price,
        total_points,
        total_net,
    )


# ============================================================
# 最適な購入先を探索
# ============================================================

def find_best(
    products,
    spu_multiplier,
    max_shops,
    min_shop_price,
    bonus_cap,
):
    """
    各商品についてAmazon / 楽天の全組み合わせを探索。
    """

    if not products:
        return None

    best = None

    for assignment_tuple in itertools.product(
        ["A", "R"],
        repeat=len(products),
    ):

        rows, buyaround = calculate_item_results(
            products=products,
            assignments=assignment_tuple,
            spu_multiplier=spu_multiplier,
            max_shops=max_shops,
            min_shop_price=min_shop_price,
            bonus_cap=bonus_cap,
        )

        total_price, total_points, total_net = (
            calculate_total(rows)
        )

        candidate = {
            "assignments": assignment_tuple,
            "rows": rows,
            "total_price": total_price,
            "total_points": total_points,
            "total_net": total_net,
            "buyaround": buyaround,
        }

        if best is None:
            best = candidate
            continue

        if candidate["total_net"] < best["total_net"]:
            best = candidate

    return best


# ============================================================
# Session State
# ============================================================

if "products" not in st.session_state:

    st.session_state.products = [
        normalize_product(DEFAULT_PRODUCT)
    ]

if "max_shops" not in st.session_state:
    st.session_state.max_shops = (
        DEFAULT_SETTINGS["max_shops"]
    )

if "min_shop_price" not in st.session_state:
    st.session_state.min_shop_price = (
        DEFAULT_SETTINGS["min_shop_price"]
    )

if "bonus_cap" not in st.session_state:
    st.session_state.bonus_cap = (
        DEFAULT_SETTINGS["bonus_cap"]
    )

if "spu_multiplier" not in st.session_state:
    st.session_state.spu_multiplier = (
        DEFAULT_SETTINGS["spu_multiplier"]
    )

if "result" not in st.session_state:
    st.session_state.result = None

if "loaded_once" not in st.session_state:
    st.session_state.loaded_once = False


# ============================================================
# 初回Google Sheets読み込み
# ============================================================

if not st.session_state.loaded_once:

    try:

        loaded = load_data_from_google_sheets()

        if loaded:
            st.session_state.loaded_from_google_sheets = True
        else:
            st.session_state.loaded_from_google_sheets = False

    except Exception:
        st.session_state.loaded_from_google_sheets = False

    st.session_state.loaded_once = True


# ============================================================
# タイトル
# ============================================================

st.title("🛒 Amazon vs 楽天 最安振り分け計算")

st.caption(
    "Amazonと楽天の価格・ポイント還元を比較して、"
    "商品ごとの最適な購入先を計算します。"
)


# ============================================================
# サイドバー
# ============================================================

with st.sidebar:

    st.header("⚙️ 設定")

    st.subheader("楽天買い回り")

    max_shops = st.number_input(
        "最大ショップ数",
        min_value=1,
        max_value=50,
        value=int(st.session_state.max_shops),
        step=1,
        key="ui_max_shops",
    )

    min_shop_price = st.number_input(
        "買い回り対象となる最低購入金額",
        min_value=0.0,
        value=float(
            st.session_state.min_shop_price
        ),
        step=100.0,
        key="ui_min_shop_price",
    )

    bonus_cap = st.number_input(
        "買い回りボーナス上限",
        min_value=0.0,
        value=float(
            st.session_state.bonus_cap
        ),
        step=100.0,
        key="ui_bonus_cap",
    )

    st.subheader("楽天SPU")

    spu_multiplier = st.number_input(
        "SPU上乗せ倍率（%）",
        min_value=0.0,
        max_value=100.0,
        value=float(
            st.session_state.spu_multiplier
        ),
        step=0.5,
        key="ui_spu_multiplier",
    )

    st.divider()

    # --------------------------------------------------------
    # 設定をSession Stateへ反映
    # --------------------------------------------------------

    st.session_state.max_shops = int(
        max_shops
    )

    st.session_state.min_shop_price = (
        min_shop_price
    )

    st.session_state.bonus_cap = (
        bonus_cap
    )

    st.session_state.spu_multiplier = (
        spu_multiplier
    )

    # --------------------------------------------------------
    # Google Sheets
    # --------------------------------------------------------

    st.subheader("💾 データ")

    if st.button(
        "💾 Google Sheetsへ保存",
        use_container_width=True,
    ):

        try:

            save_data_to_google_sheets()

            st.success(
                "Google Sheetsへ保存しました。"
            )

        except Exception as e:

            st.error(
                f"保存に失敗しました。\n\n{e}"
            )

    if st.button(
        "📥 Google Sheetsから読み込み",
        use_container_width=True,
    ):

        try:

            loaded = load_data_from_google_sheets()

            if loaded:

                st.success(
                    "Google Sheetsから読み込みました。"
                )

                st.rerun()

            else:

                st.error(
                    "Google Sheetsから読み込めませんでした。"
                )

        except Exception as e:

            st.error(
                f"読み込みに失敗しました。\n\n{e}"
            )


# ============================================================
# 商品入力
# ============================================================

st.header("📦 商品")

st.caption(
    "Amazon価格・Amazon還元率・楽天価格・楽天還元率を入力してください。"
    "楽天URLを入力すると楽天APIから価格・還元率を取得できます。"
)


# ------------------------------------------------------------
# 商品追加
# ------------------------------------------------------------

if st.button(
    "＋ 商品を追加",
    use_container_width=False,
):

    st.session_state.products.append(
        normalize_product(DEFAULT_PRODUCT)
    )

    st.rerun()


# ============================================================
# 商品入力フォーム
# ============================================================

delete_index = None

for index, product in enumerate(
    st.session_state.products
):

    with st.container(border=True):

        col_title, col_delete = st.columns(
            [8, 1]
        )

        with col_title:

            st.markdown(
                f"### 商品 {index + 1}"
            )

        with col_delete:

            if st.button(
                "🗑️",
                key=f"delete_{index}",
                help="この商品を削除",
            ):

                delete_index = index

        col1, col2 = st.columns(2)

        with col1:

            name = st.text_input(
                "商品名",
                value=product["name"],
                key=f"name_{index}",
            )

            ap = st.number_input(
                "Amazon価格",
                min_value=0.0,
                value=float(product["ap"]),
                step=1.0,
                key=f"ap_{index}",
            )

            apt = st.number_input(
                "Amazon還元率（%）",
                min_value=0.0,
                max_value=100.0,
                value=float(product["apt"]),
                step=0.1,
                key=f"apt_{index}",
            )

            baby = st.checkbox(
                "らくベビ割を適用",
                value=bool(product["baby"]),
                key=f"baby_{index}",
            )

        with col2:

            rurl = st.text_input(
                "楽天商品URL",
                value=product["rurl"],
                key=f"rurl_{index}",
            )

            rp = st.number_input(
                "楽天価格",
                min_value=0.0,
                value=float(product["rp"]),
                step=1.0,
                key=f"rp_{index}",
            )

            rpt = st.number_input(
                "楽天還元倍率（%）",
                min_value=0.0,
                max_value=100.0,
                value=float(product["rpt"]),
                step=0.1,
                key=f"rpt_{index}",
            )

            if rurl:

                if st.button(
                    "🔎 楽天APIから取得",
                    key=f"rakuten_{index}",
                ):

                    with st.spinner(
                        "楽天の商品情報を取得しています..."
                    ):

                        try:

                            result = (
                                fetch_rakuten_price_and_point(
                                    url=rurl,
                                    keyword=name,
                                )
                            )

                            if result:

                                st.session_state.products[
                                    index
                                ]["rp"] = result[
                                    "price"
                                ]

                                st.session_state.products[
                                    index
                                ]["rpt"] = result[
                                    "point_rate"
                                ]

                                if (
                                    not name
                                    and result["item_name"]
                                ):
                                    st.session_state.products[
                                        index
                                    ]["name"] = result[
                                        "item_name"
                                    ]

                                st.success(
                                    "楽天情報を取得しました。"
                                )

                                st.rerun()

                            else:

                                st.warning(
                                    "楽天の商品情報を取得できませんでした。"
                                )

                        except Exception as e:

                            st.error(
                                f"楽天APIエラー:\n{e}"
                            )

        # ----------------------------------------------------
        # Session Stateへ保存
        # ----------------------------------------------------

        st.session_state.products[index] = {
            "name": name,
            "ap": ap,
            "apt": apt,
            "baby": baby,
            "rp": rp,
            "rpt": rpt,
            "rurl": rurl,
        }


# ============================================================
# 商品削除
# ============================================================

if delete_index is not None:

    if len(st.session_state.products) > 1:

        st.session_state.products.pop(
            delete_index
        )

    else:

        st.session_state.products = [
            normalize_product(DEFAULT_PRODUCT)
        ]

    st.rerun()


# ============================================================
# 計算ボタン
# ============================================================

st.divider()

if st.button(
    "🚀 最安購入先を計算",
    type="primary",
    use_container_width=True,
):

    # 空の商品を除外
    products = []

    for product in st.session_state.products:

        normalized = normalize_product(
            product
        )

        if (
            normalized["name"].strip()
            or normalized["ap"] > 0
            or normalized["rp"] > 0
        ):
            products.append(normalized)

    if not products:

        st.warning(
            "商品を1つ以上入力してください。"
        )

    else:

        with st.spinner(
            "最安購入先を計算しています..."
        ):

            result = find_best(
                products=products,
                spu_multiplier=(
                    st.session_state.spu_multiplier
                ),
                max_shops=(
                    st.session_state.max_shops
                ),
                min_shop_price=(
                    st.session_state.min_shop_price
                ),
                bonus_cap=(
                    st.session_state.bonus_cap
                ),
            )

            st.session_state.result = result


# ============================================================
# 計算結果
# ============================================================

if st.session_state.result:

    result = st.session_state.result

    st.divider()

    st.header("🏆 最適購入結果")

    # --------------------------------------------------------
    # 合計
    # --------------------------------------------------------

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "購入金額",
            yen(result["total_price"]),
        )

    with col2:

        st.metric(
            "還元ポイント",
            yen(result["total_points"]),
        )

    with col3:

        st.metric(
            "実質負担額",
            yen(result["total_net"]),
        )

    # --------------------------------------------------------
    # 買い回り
    # --------------------------------------------------------

    buyaround = result["buyaround"]

    st.subheader("🛍️ 楽天買い回り")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "ショップ数",
            f"{buyaround['shop_count']}店舗",
        )

    with col2:

        st.metric(
            "買い回り上乗せ",
            f"{buyaround['bonus_multiplier']}%",
        )

    with col3:

        st.metric(
            "買い回りポイント",
            yen(buyaround["bonus_points"]),
        )

    # --------------------------------------------------------
    # 商品別結果
    # --------------------------------------------------------

    st.subheader("📋 商品別結果")

    display_rows = []

    for row in result["rows"]:

        display_rows.append({
            "商品名": row["商品名"],
            "購入先": row["購入先"],
            "価格": yen(row["価格"]),
            "還元倍率": (
                f"{row['還元倍率']:.1f}%"
            ),
            "還元ポイント": yen(
                row["還元ポイント"]
            ),
            "実質負担額": yen(
                row["実質負担額"]
            ),
        })

    st.table(
        display_rows
    )

    # --------------------------------------------------------
    # 結果の詳細
    # --------------------------------------------------------

    st.subheader("💰 購入先の内訳")

    amazon_count = sum(
        1
        for row in result["rows"]
        if row["購入先"] == "Amazon"
    )

    rakuten_count = sum(
        1
        for row in result["rows"]
        if row["購入先"] == "楽天"
    )

    col1, col2 = st.columns(2)

    with col1:

        st.info(
            f"Amazon：{amazon_count}商品"
        )

    with col2:

        st.info(
            f"楽天：{rakuten_count}商品"
        )


# ============================================================
# API設定確認
# ============================================================

with st.expander(
    "🔧 API / Google Sheets設定の状態"
):

    app_id, access_key, referer = (
        get_rakuten_credentials()
    )

    if app_id:
        st.success(
            "楽天 Application ID：設定済み"
        )
    else:
        st.error(
            "楽天 Application ID：未設定"
        )

    if access_key:
        st.success(
            "楽天 Access Key：設定済み"
        )
    else:
        st.error(
            "楽天 Access Key：未設定"
        )

    if referer:
        st.success(
            "楽天 Referer：設定済み"
        )
    else:
        st.warning(
            "楽天 Referer：未設定"
        )

    try:

        sheet_id = get_secret(
            "GOOGLE_SHEET_ID"
        )

        if sheet_id:
            st.success(
                "Google Sheets ID：設定済み"
            )
        else:
            st.error(
                "Google Sheets ID：未設定"
            )

    except Exception:

        st.error(
            "Google Sheets ID：確認できません"
        )


# ============================================================
# フッター
# ============================================================

st.divider()

st.caption(
    "Amazon・楽天の価格やポイント条件は変動するため、"
    "最終的な購入前に各販売サイトでご確認ください。"
)
