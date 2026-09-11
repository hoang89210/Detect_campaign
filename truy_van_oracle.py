import oracledb
from sqlalchemy import create_engine, text
import pandas as pd
import re

from datetime import datetime

# Hàm làm sạch các ký tự không hợp lệ trong Excel
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
def clean_excel_value(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value

# --- 1. THÔNG TIN CẤU HÌNH ---
config = {
    "user": "HOANGTT",
    "password": "NRYVvpRnOr3XpV4VEVyr",
    "dsn": "172.25.80.171:11521/EDUCA" # Ví dụ: "192.168.1.10:1521/orcl"
}





# Đường dẫn file Excel đầu ra
file_name = f"Data_Campaign_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# --- 2. CÂU TRUY VẤN SQL ---
sql_query = r"""   
WITH t100 AS (
                SELECT 
                    SHIPMENT_DATA_TYPE,
                    CRM3_CAMPAIGN_ID,
                    DEAL_SHIPMENT_ID,
                    CRM1_ORDER_ID,
                    CREATED_BY_ID,
                    DEAL_SALES_ID,
                    DEAL_SALES_CREATED_DATE,
                    DEAL_SALES_ASSIGNED_DATE,
                    DEAL_SHIPMENT_STAGE,
                    ORDER_FREEZE_STATUS,
                    ORDER_CURRENT_STATUS,
                    SHIPMENT_ASSIGNED_BY_ID,
                    SHIPMENT_SOURCE_ID,
                    SHIPMENT_CREATED_DATE,
                    ORDER_DATE,
                    CREATED_DATE,
                    CREATED_AT,
                    CURRENT_ORDER_DATE,
                    SHIPMENT_MODIFIED_DATE,
                    REVENUE_DEPARTMENT_ID,
                    SHIPMENT_CATEGORY_ID,
                    CONTACT_ID,
                    GROUP_STATUS,
                    CLOSE_DEAL_DATE,
                    SHIPMENT_SALES_DEPARTMENT,
                    FREEZE_PRICE,
                    CRM3_OPPORTUNITY,
                    DEAL_PIPELINE_ID,
                    WON_IN_DAY,
                    TABLE_SOURCE,
                    CAMPAIGN_NAME,
                    SHIP_CODE,
                    PHONE,
                    CURRENT_PRICE,
                    CRM1_DEPARTMENT_ID,
                    USERNAME,
                    CRM3_SALES_ID,
                    DEAL_SALES_DISTRIBUTOR,
                    IS_WON_BY_FIRST_ASSIGN,
                    IS_UNIQUE_ORDER_ID,
                    CRM1_PACKAGE_ID,
                    IS_CONTACT_TEST,
                    CRM1_UPDATED_AT
                FROM ODS_DA.FACT_CRM_SHIPMENT_FREEZE

                UNION ALL

                SELECT 
                    SHIPMENT_DATA_TYPE,
                    CRM3_CAMPAIGN_ID,
                    DEAL_SHIPMENT_ID,
                    CRM1_ORDER_ID,
                    CREATED_BY_ID,
                    DEAL_SALES_ID,
                    DEAL_SALES_CREATED_DATE,
                    DEAL_SALES_ASSIGNED_DATE,
                    DEAL_SHIPMENT_STAGE,
                    ORDER_FREEZE_STATUS,
                    ORDER_CURRENT_STATUS,
                    SHIPMENT_ASSIGNED_BY_ID,
                    SHIPMENT_SOURCE_ID,
                    SHIPMENT_CREATED_DATE,
                    ORDER_DATE,
                    CREATED_DATE,
                    CREATED_AT,
                    CURRENT_ORDER_DATE,
                    SHIPMENT_MODIFIED_DATE,
                    REVENUE_DEPARTMENT_ID,
                    SHIPMENT_CATEGORY_ID,
                    CONTACT_ID,
                    GROUP_STATUS,
                    CLOSE_DEAL_DATE,
                    SHIPMENT_SALES_DEPARTMENT,
                    FREEZE_PRICE,
                    CRM3_OPPORTUNITY,
                    DEAL_PIPELINE_ID,
                    WON_IN_DAY,
                    TABLE_SOURCE,
                    CAMPAIGN_NAME,
                    SHIP_CODE,
                    PHONE,
                    CURRENT_PRICE,
                    CRM1_DEPARTMENT_ID,
                    USERNAME,
                    CRM3_SALES_ID,
                    DEAL_SALES_DISTRIBUTOR,
                    IS_WON_BY_FIRST_ASSIGN,
                    IS_UNIQUE_ORDER_ID,
                    CRM1_PACKAGE_ID,
                    IS_CONTACT_TEST,
                    CRM1_UPDATED_AT
                FROM ODS_DA.FACT_CRM_SHIPMENT_TODAY_INPUT
            ),
            rp AS (
            SELECT 
                t100.CRM3_CAMPAIGN_ID,
                t100.DEAL_SHIPMENT_STAGE,
                t100.SHIPMENT_ASSIGNED_BY_ID,
                t100.SHIPMENT_SOURCE_ID,
                t100.SHIPMENT_CREATED_DATE,
                t100.CREATED_DATE,
                t100.SHIPMENT_MODIFIED_DATE,
                TRUNC(t100.SHIPMENT_MODIFIED_DATE) AS MODIFIED_DATE,
                t100.REVENUE_DEPARTMENT_ID,
                t100.SHIPMENT_CATEGORY_ID,
                t100.CONTACT_ID,
                t100.GROUP_STATUS,
                t100.CLOSE_DEAL_DATE,
                t100.SHIPMENT_SALES_DEPARTMENT,
                t100.FREEZE_PRICE,
                t100.CREATED_BY_ID,
                t100.WON_IN_DAY,
                t100.DEAL_SALES_ID,
                t100.DEAL_SALES_CREATED_DATE,
                t100.DEAL_SALES_ASSIGNED_DATE,
                t100.CRM3_OPPORTUNITY,
                t100.DEAL_PIPELINE_ID,
                t100.ORDER_DATE,
                t100.TABLE_SOURCE,
                t100.CAMPAIGN_NAME,
                t100.DEAL_SHIPMENT_ID,
                t100.CRM1_ORDER_ID,
                t100.ORDER_FREEZE_STATUS,
                t100.ORDER_CURRENT_STATUS,
                t100.SHIP_CODE,
                t100.PHONE,
                t100.CURRENT_PRICE,
                t100.CURRENT_ORDER_DATE,
                t100.USERNAME,
                t100.CRM1_DEPARTMENT_ID,
                t100.SHIPMENT_DATA_TYPE,
                t100.CRM1_PACKAGE_ID,
                t100.IS_CONTACT_TEST,
                CASE 
                    WHEN t100.ORDER_DATE = t100.CURRENT_ORDER_DATE THEN '1.Đóng băng unique' 
                    ELSE '2.Đóng băng trùng' 
                END AS ORDER_FREEZE_DUPLICATE_STATUS,
                t1.GROUP_ID,
                deal.UF_CRM_UTM_CONTENT,
                deal.VALUE_ID,
             	bu.login 
            FROM t100
            LEFT JOIN EDUCA_CRM.CRM_ORDER t1 
                ON t100.CRM1_ORDER_ID = t1.ID
            LEFT JOIN ODS_DA.BOD_DIM_CAMPAIGNS dimcamp 
                ON t100.CRM3_CAMPAIGN_ID = dimcamp.ID
            LEFT JOIN BITRIX24.B_UTS_CRM_DEAL deal 
                ON t100.DEAL_SHIPMENT_ID = deal.UF_CRM_LINK_DEAL_SHIPMENT
            LEFT JOIN BITRIX24.B_UTS_CRM_LEAD ls ON ls.UF_CRM_LINK_NEW_DEAL = deal.value_id
            LEFT JOIN BITRIX24.B_CRM_LEAD l1 ON l1.id = ls.VALUE_ID 
            LEFT JOIN  BITRIX24.B_USER bu ON bu.ID = l1.ASSIGNED_BY_ID 
            WHERE t100.FREEZE_PRICE > 0
              AND dimcamp.CAMP_LEVEL2 = '1.3 Uniclass - DEMO CTV'
              AND  dimcamp.campaign_name ='Inhouse_EdupiaClass_TU_Demo_Team_C_c1')
              
              SELECT phone, current_order_date FROM rp
		
"""

def export_data():
    conn = None
    try:
        print(" đang kết nối tới Oracle server...")
        # Kết nối tới Oracle
        conn = oracledb.connect(
            user=config["user"],
            password=config["password"],
            dsn=config["dsn"]
        )
        
        print(" Chạy truy vấn SQL... (Vui lòng đợi)")
        # Sử dụng pandas để đọc dữ liệu trực tiếp vào DataFrame
        df = pd.read_sql(sql_query, conn)
        
        if df.empty:
            print(" Không có dữ liệu trả về từ truy vấn.")
        else:
            print(f" Đã lấy được {len(df)} dòng. Đang xuất file Excel...")
            # Làm sạch dữ liệu trước khi xuất Excel
            df = df.map(clean_excel_value)
            # Xuất ra file Excel
            df.to_excel(file_name, index=False, engine='openpyxl')
            print(f" Thành công! File đã được lưu tại: {file_name}")

    except Exception as e:
        print(f" Lỗi xảy ra: {e}")
        
    finally:
        # Đóng kết nối
        if conn:
            conn.close()
            print(" Đã đóng kết nối Oracle.")

if __name__ == "__main__":
    export_data()