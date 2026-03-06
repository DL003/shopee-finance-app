import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (全局千位符修复版)")

# --- 核心辅助：清洗印尼/特殊格式金额 ---
def clean_currency(value):
    """
    深度清洗金额：将 '125.795' 还原为 125795.0
    """
    if pd.isna(value) or str(value).strip() == "" or str(value).strip() == "-":
        return 0.0
    
    val_str = str(value).strip()
    # 1. 移除千分位点
    val_str = val_str.replace('.', '')
    # 2. 处理可能存在的逗号小数点
    val_str = val_str.replace(',', '.')
    # 3. 移除非数字字符 (保留负号和点)
    val_str = re.sub(r'[^0-9.\-]', '', val_str)
    
    try:
        return float(val_str)
    except:
        return 0.0

def get_col_exact(df, target_name):
    target_clean = str(target_name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target_clean:
            return col
    for col in df.columns:
        if target_clean in str(col).strip().lower():
            return col
    return None

# 文件上传
col1, col2 = st.columns(2)
with col1:
    f_a = st.file_uploader("1. 上传【模板表.xlsx】", type=['xlsx', 'csv'])
    f_b = st.file_uploader("2. 上传【订单表.xlsx】", type=['xlsx', 'csv'])
with col2:
    f_c = st.file_uploader("3. 上传【订单收入.xlsx】", type=['xlsx', 'csv'])
    f_d = st.file_uploader("4. 上传【成本表.xlsx】", type=['xlsx', 'csv'])

if f_a and f_b and f_c and f_d:
    if st.button("🚀 开始全局修复并核算"):
        try:
            def load_df(file):
                try: return pd.read_excel(file, dtype=str)
                except: return pd.read_csv(file, dtype=str)

            df_temp = load_df(f_a)
            df_order = load_df(f_b)
            df_income = load_df(f_c)
            df_cost = load_df(f_d)

            # --- 关键步骤：全局清洗所有表的金额列 ---
            # 定义所有需要清洗的金额关键字
            money_keywords = ['harga', 'diskon', 'voucher', 'fee', 'commission', 'premium', 'service', 'processing', '成本', '单价', 'total']
            
            def clean_df_money(df):
                for col in df.columns:
                    col_lower = str(col).lower()
                    if any(k in col_lower for k in money_keywords):
                        df[col] = df[col].apply(clean_currency)
                return df

            df_order = clean_df_money(df_order)
            df_income = clean_df_money(df_income)
            df_cost = clean_df_money(df_cost)

            # 寻找核心关联键
            oid_order = get_col_exact(df_order, 'order number')
            oid_income = get_col_exact(df_income, 'Order number')
            sku_order = get_col_exact(df_order, 'Nomor Referensi SKU')
            sku_cost = get_col_exact(df_cost, 'Nomor Referensi SKU')
            cost_val_col = get_col_exact(df_cost, '成本单价')

            if not oid_order or not oid_income:
                st.error("❌ 无法匹配订单号列。")
                st.stop()

            # 1. 预处理收入表费用明细 (此时数据已清洗为浮点数)
            income_fee_map = {
                'ams': 'AMS Commission Fee',
                'comm': 'Commission fee (including PPN 10%)',
                'service': 'Service Fee',
                'proc': 'Seller Order Processing Fee',
                'premium': 'Premium'
            }
            
            df_income_final = df_income.drop_duplicates(oid_income)

            # 2. 预处理订单表
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')
            # 数量确保为数值
            df_order[qty_col] = pd.to_numeric(df_order[qty_col], errors='coerce').fillna(0)

            # 3. 关联数据
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            df_main = pd.merge(df_order, df_income_final, left_on=oid_order, right_on=oid_income, how='left')
            
            # 关联成本
            if sku_order and sku_cost and cost_val_col:
                df_main = pd.merge(df_main, df_cost[[sku_cost, cost_val_col]], left_on=sku_order, right_on=sku_cost, how='left')
            else:
                df_main['成本单价'] = 0.0

            # 4. 财务核算逻辑 (包含分摊)
            status_col = get_col_exact(df_order, 'Status Pesanan')
            
            def calc_row(row):
                st_str = str(row.get(status_col, '')).lower()
                if 'batal' in st_str or 'cancel' in st_str:
                    return pd.Series([0.0]*9)
                
                cnt = row['计数'] if row['计数'] > 0 else 1
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / cnt
                
                # 提取并平摊费用
                def get_fee(key):
                    col = get_col_exact(df_income, income_fee_map[key])
                    return row.get(col, 0) / cnt if col else 0
                
                f_ams = get_fee('ams')
                f_comm = get_fee('comm')
                f_service = get_fee('service')
                f_proc = get_fee('proc')
                f_prem = get_fee('premium')
                
                inc = s_amt - v_share + (f_ams + f_comm + f_service + f_proc + f_prem)
                c_total = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return pd.Series([s_amt, v_share, inc, c_total, f_ams, f_comm, f_service, f_proc, f_prem])

            calc_res_cols = ['_S', '_V', '_I', '_C', '_f_ams', '_f_comm', '_f_service', '_f_proc', '_f_prem']
            df_main[calc_res_cols] = df_main.apply(calc_row, axis=1)

            # 5. 填充回模板 A
            df_final = pd.DataFrame(columns=df_temp.columns)
            
            special_map = {
                '成功订单销售金额': '_S',
                '优惠券': '_V',
                'income': '_I',
                '成本': cost_val_col,
                '总成本': '_C'
            }

            for col in df_final.columns:
                c_name = str(col).strip()
                
                # A. 费用明细 (模糊匹配关键字并填入平摊后的值)
                matched_fee = False
                for k_internal, k_orig in income_fee_map.items():
                    if k_orig.lower() in c_name.lower():
                        df_final[col] = df_main[f'_f_{k_internal}']
                        matched_fee = True
                        break
                if matched_fee: continue

                # B. 特殊计算列
                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    # C. 原始字段 (此时 df_main 里的金额已全部清洗为数字)
                    orig = get_col_exact(df_main, c_name)
                    if orig: df_final[col] = df_main[orig]

            st.success("✅ 全局金额格式修复完成！所有点号已处理。")
            st.dataframe(df_final.head(10))

            output = io.BytesIO()
            df_final.to_excel(output, index=False)
            st.download_button("📥 下载全局修复后的表A", output.getvalue(), "Shopee_Global_Fixed_A.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
