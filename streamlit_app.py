import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (千位符修复版)")

# --- 核心辅助：清洗印尼格式金额 ---
def clean_currency(value):
    """
    将 '299.000' 或 '299,000' 转换为 299000
    """
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    
    # 移除所有非数字和非负号字符 (保留负号以防费用项)
    # 但由于千位符是点，我们先移除所有的点和逗号
    val_str = str(value).replace('.', '').replace(',', '').strip()
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
    if st.button("🚀 开始修复金额并核算"):
        try:
            def load_df(file):
                try: return pd.read_excel(file)
                except: return pd.read_csv(file)

            df_temp = load_df(f_a)
            df_order = load_df(f_b)
            df_income = load_df(f_c)
            df_cost = load_df(f_d)

            # 寻找核心关联键
            oid_order = get_col_exact(df_order, 'order number')
            oid_income = get_col_exact(df_income, 'Order number')
            sku_order = get_col_exact(df_order, 'Nomor Referensi SKU')
            sku_cost = get_col_exact(df_cost, 'Nomor Referensi SKU')
            cost_val_col = get_col_exact(df_cost, '成本单价')

            if not oid_order or not oid_income:
                st.error("❌ 无法匹配订单号列。")
                st.stop()

            # 1. 预处理收入表 (清洗金额)
            income_fee_map = {
                'ams': 'AMS Commission Fee',
                'comm': 'Commission fee (including PPN 10%)',
                'service': 'Service Fee',
                'proc': 'Seller Order Processing Fee',
                'premium': 'Premium'
            }
            
            df_income_clean = df_income[[oid_income]].copy()
            for key, target in income_fee_map.items():
                actual_col = get_col_exact(df_income, target)
                if actual_col:
                    # 使用 clean_currency 处理费用列
                    df_income_clean[key] = df_income[actual_col].apply(clean_currency)
                else:
                    df_income_clean[key] = 0.0
            
            df_income_final = df_income_clean.drop_duplicates(oid_income)

            # 2. 预处理订单表 (清洗金额)
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')
            
            df_order[price_col] = df_order[price_col].apply(clean_currency)
            df_order[voucher_col] = df_order[voucher_col].apply(clean_currency)
            # 数量通常是整数，直接转换
            df_order[qty_col] = pd.to_numeric(df_order[qty_col], errors='coerce').fillna(0)

            # 3. 关联数据
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            df_main = pd.merge(df_order, df_income_final, left_on=oid_order, right_on=oid_income, how='left')
            
            # 关联成本并清洗
            if sku_order and sku_cost and cost_val_col:
                df_cost_sub = df_cost[[sku_cost, cost_val_col]].copy()
                df_cost_sub[cost_val_col] = df_cost_sub[cost_val_col].apply(clean_currency)
                df_main = pd.merge(df_main, df_cost_sub, left_on=sku_order, right_on=sku_cost, how='left')
            else:
                df_main['成本单价'] = 0.0

            # 4. 财务核算
            status_col = get_col_exact(df_order, 'Status Pesanan')
            
            def calc_row(row):
                st_str = str(row.get(status_col, '')).lower()
                if 'batal' in st_str or 'cancel' in st_str:
                    return 0.0, 0.0, 0.0, 0.0
                
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                fees = row['ams'] + row['comm'] + row['service'] + row['proc'] + row['premium']
                
                # 现在计算出来的 inc 应该是正值了
                inc = s_amt - v_share + fees
                c_total = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                return s_amt, v_share, inc, c_total

            df_main[['_S', '_V', '_I', '_C']] = df_main.apply(lambda r: pd.Series(calc_row(r)), axis=1)

            # 5. 填充回模板
            df_final = pd.DataFrame(columns=df_temp.columns)
            special_map = {'成功订单销售金额': '_S', '优惠券': '_V', 'income': '_I', '成本': cost_val_col, '总成本': '_C'}

            for col in df_final.columns:
                c_name = str(col).strip()
                # 费用明细
                matched_fee = False
                for k_internal, k_orig in income_fee_map.items():
                    if k_orig.lower() in c_name.lower():
                        df_final[col] = df_main[k_internal]
                        matched_fee = True
                        break
                if matched_fee: continue

                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    orig = get_col_exact(df_main, c_name)
                    if orig: df_final[col] = df_main[orig]

            st.success("✅ 金额格式已修正，核算完成！")
            st.dataframe(final_show := df_final.head(10))

            output = io.BytesIO()
            df_final.to_excel(output, index=False)
            st.download_button("📥 下载 Shopee 核算结果", output.getvalue(), "Shopee_Corrected_A.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
