import oracledb
import pandas as pd
import re
import sys
from datetime import datetime

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
def clean_excel_value(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

config = {
    "user": "HOANGTT",
    "password": "NRYVvpRnOr3XpV4VEVyr",
    "dsn": "172.25.80.171:11521/EDUCA"
}

LI_DO_TUTOR_SQL = """
WITH dim_stage AS (
    SELECT id, CATEGORY_ID, NAME STAGE_NAME, STATUS_ID, SORT
    FROM BITRIX24.B_CRM_STATUS
    WHERE CATEGORY_ID = 11 AND ENTITY_ID LIKE 'DEAL_STAGE%%'
), base_data AS (
    SELECT
        deal.id,
        MAPPHONE.VALUE,
        deal.STAGE_ID,
        camp.CAMP_LEVEL2,
        MIN(CASE WHEN dim_stage.sort >= 30 THEN stage.CREATED_TIME END) AS l1,
        MIN(CASE WHEN dim_stage.sort >= 40 THEN stage.CREATED_TIME END) AS l2,
        MIN(CASE WHEN dim_stage.sort >= 50 THEN stage.CREATED_TIME END) AS l3,
        MIN(CASE WHEN dim_stage.sort >= 60 THEN stage.CREATED_TIME END) AS l4,
        MIN(CASE WHEN dim_stage.sort >= 70 THEN stage.CREATED_TIME END) AS l5,
        MIN(CASE WHEN dim_stage.sort >= 80 THEN stage.CREATED_TIME END) AS l6,
        MIN(CASE WHEN dim_stage.sort >= 90 THEN stage.CREATED_TIME END) AS l7,
        MIN(CASE WHEN dim_stage.sort >= 100 THEN stage.CREATED_TIME END) AS l8,
        TRUNC(deal.DATE_CREATE) AS date_create,
        max(dim_stage.SORT) as max_deal_stage_sort
    FROM BITRIX24.B_CRM_DEAL deal
    LEFT JOIN BITRIX24.B_UTS_CRM_DEAL detail_deal ON detail_deal.VALUE_ID = deal.ID
    LEFT JOIN BITRIX24.B_CRM_FIELD_MULTI MAPPHONE
        ON deal.CONTACT_ID = MAPPHONE.ELEMENT_ID
       AND MAPPHONE.ENTITY_ID = 'CONTACT'
       AND MAPPHONE.TYPE_ID = 'PHONE'
    LEFT JOIN BITRIX24.B_CRM_DEAL_STAGE_HISTORY stage
        ON stage.OWNER_ID = deal.ID
    LEFT JOIN dim_stage
        ON stage.STAGE_ID = dim_stage.STATUS_ID
       AND deal.CATEGORY_ID = dim_stage.CATEGORY_ID
    LEFT JOIN ods_da.BOD_DIM_CAMPAIGNS camp
        ON detail_deal.UF_CRM_CAMPAIGN_CODE = camp.ID
    WHERE dim_stage.SORT >= 20 and dim_stage.SORT <= 100
    GROUP BY deal.id, MAPPHONE.VALUE, deal.STAGE_ID, camp.CAMP_LEVEL2, TRUNC(deal.DATE_CREATE)
),
li_do_fail AS (
    SELECT *
    FROM (
        SELECT
            base_data.*,
            ROW_NUMBER() OVER (PARTITION BY VALUE ORDER BY date_create DESC) AS rnk,
            dim_stage.STAGE_NAME
        FROM base_data
        LEFT JOIN dim_stage ON dim_stage.sort = base_data.max_deal_stage_sort
    )
    WHERE rnk = 1
)
SELECT value, stage_name FROM li_do_fail WHERE VALUE IN ({placeholders})
"""

CHECK_MUA_SQL = """
WITH b AS (
    SELECT
        extract(MONTH FROM orders.ORDER_DATE) AS m,
        ORDER_DATE AS ngay_ban,
        ORDERS.PHONE phone,
        orders.CAMPAIGN_NAME,
        crm_pack.name AS pack,
        orders.CUSTOMER_NAME,
        orders.PRICE,
        cc.PRODUCT,
        cc.PRODUCT_GROUP,
        cc.PRODUCT_NAME_SMS,
        orders.ORDER_DATE,
        orders.CRM_CONTACT_ID,
        orders.CRM_CHILDREN_ID,
        dense_rank() over(PARTITION BY phone ORDER BY product_group asc) AS lan_mua,
        row_number() over(PARTITION BY phone ORDER BY order_date desc) AS rnk
    FROM EDUCA_CRM.CRM_ORDER orders
    LEFT JOIN EDUCA_CRM.CRM_Package crm_pack ON ORDERS.PACKAGE_ID = CRM_PACK.ID
    LEFT JOIN educa_crm.CRM_CATEGORIES cc ON cc.id = crm_pack."TYPE"
    LEFT JOIN ODS_DA.BOD_DIM_CAMPAIGNS bdc ON bdc.CAMPAIGN_NAME = orders.CAMPAIGN_NAME
    WHERE orders.price > 0 AND cc.STATUS = 1
    AND crm_pack.NAME NOT LIKE '%%uni%%'
)
SELECT lan_mua, phone, PRODUCT_GROUP FROM b WHERE phone IN ({placeholders}) AND rnk = 1
"""


def classify_stage(stage_name):
    if stage_name is None or (isinstance(stage_name, float) and pd.isna(stage_name)):
        return "Lead"
    s = str(stage_name).strip()
    if any(s.startswith(f"L{i}") for i in range(1, 5)):
        return "Lead"
    if any(s.startswith(f"L{i}") for i in range(5, 9)):
        return "Deal"
    return "Lead"


def classify_isbuy(lan_mua, product_group):
    if lan_mua is None or (isinstance(lan_mua, float) and pd.isna(lan_mua)):
        return "nobuy", None
    if lan_mua >= 2:
        return "isbuy_2sp", None
    if lan_mua == 1:
        prod = str(product_group).strip() if product_group is not None and not (isinstance(product_group, float) and pd.isna(product_group)) else None
        return "isbuy", prod
    return "nobuy", None


def map_camp(stage_type, isbuy_type, product, cap):
    product_lower = {
        "edupia": "Edu", "edu": "Edu",
        "babilala": "Babi", "babi": "Babi",
        "tutor": "Tutor",
    }
    prod_key = None
    if product is not None and not (isinstance(product, float) and pd.isna(product)):
        prod_lower = str(product).strip().lower()
        for k, v in product_lower.items():
            if k in prod_lower:
                prod_key = v
                break

    cap_suffix = f"_c{cap}"

    if isbuy_type == "nobuy":
        return f"Inhouse_EdupiaClass_Tu-van_Demo_nobuy_{stage_type}{cap_suffix}"
    if isbuy_type == "isbuy_2sp":
        return f"Inhouse_EdupiaClass_Tu-van_Demo_isbuy_{stage_type}_2sp{cap_suffix}"
    if isbuy_type == "isbuy" and prod_key:
        if stage_type == "Lead":
            return f"Inhouse_EdupiaClass_Tu-van-Demo_isbuy_{prod_key}_lead{cap_suffix}"
        else:
            return f"Inhouse_EdupiaClass_Tu-van-Demo_isbuy_{prod_key}_deal{cap_suffix}"
    return f"Inhouse_EdupiaClass_Tu-van_Demo_nobuy_{stage_type}{cap_suffix}"


def run_query(conn, sql, phones):
    placeholders = ",".join([f"'{str(p)}'" for p in phones])
    final_sql = sql.format(placeholders=placeholders)
    return pd.read_sql(final_sql, conn)


def main():
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = r"C:\Users\Administrator\Desktop\Anal_C_Cuc\phone_can_camp.xlsx"

    print(f"Doc file: {input_file}")

    if input_file.endswith('.ods'):
        df_input = pd.read_excel(input_file, engine='odf')
    else:
        df_input = pd.read_excel(input_file)

    phones = df_input["phone"].astype(str).tolist()

    print(f"Tong so phone: {len(phones)}")

    conn = oracledb.connect(
        user=config["user"],
        password=config["password"],
        dsn=config["dsn"]
    )

    print("Dang query Li_do_tutor...")
    df_stage = run_query(conn, LI_DO_TUTOR_SQL, phones)
    print(f"  lay duoc {len(df_stage)} ban ghi")

    print("Dang query Check_mua...")
    df_buy = run_query(conn, CHECK_MUA_SQL, phones)
    print(f"  lay duoc {len(df_buy)} ban ghi")

    conn.close()

    print("  Li_do_tutor columns:", df_stage.columns.tolist())
    print("  Check_mua columns:", df_buy.columns.tolist())
    df_stage.columns = [c.lower() for c in df_stage.columns]
    df_buy.columns = [c.lower() for c in df_buy.columns]
    df_stage["phone"] = df_stage["value"].astype(str)
    df_stage = df_stage[["phone", "stage_name"]].drop_duplicates(subset=["phone"])

    df_buy["phone"] = df_buy["phone"].astype(str)
    df_buy = df_buy[["phone", "lan_mua", "product_group"]].drop_duplicates(subset=["phone"])

    df_input["phone"] = df_input["phone"].astype(str)

    result = df_input[["phone", "cap"]].copy()
    result = result.merge(df_stage, on="phone", how="left")
    result = result.merge(df_buy, on="phone", how="left")

    result["stage"] = result["stage_name"].apply(classify_stage)

    isbuy_data = result.apply(lambda row: classify_isbuy(row["lan_mua"], row["product_group"]), axis=1)
    result["isbuy_type"] = [x[0] for x in isbuy_data]
    result["product_resolved"] = [x[1] for x in isbuy_data]

    result["CAMPAIGN_MAPPED"] = result.apply(
        lambda row: map_camp(row["stage"], row["isbuy_type"], row["product_resolved"], row["cap"]),
        axis=1
    )

    output = result[["phone", "cap", "stage_name", "lan_mua", "product_group", "CAMPAIGN_MAPPED"]].copy()
    output.columns = ["phone", "cap", "stage_name", "lan_mua", "product_group", "CAMPAIGN_MAPPED"]

    output = output.map(clean_excel_value)

    base_name = input_file.rsplit('.', 1)[0]
    out_file = f"{base_name}_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output.to_excel(out_file, index=False, engine="openpyxl")
    print(f"\nDa xuat file: {out_file}")
    print(f"Tong: {len(output)} dong")

    print("\nPhan loai stage:")
    print(result["stage"].value_counts().to_string())
    print("\nPhan loai isbuy:")
    print(result["isbuy_type"].value_counts().to_string())
    print("\nSample 10 dong dau:")
    print(output.head(10).to_string())


if __name__ == "__main__":
    main()
