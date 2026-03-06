import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (报错修复加固版)")

# --- 核心辅助：清洗印尼/特殊格式金额 ---
def clean_currency(value):
    """
    深度清洗金额：将 '125.795' 还原为 125795.0
    """
    if pd.isna(value) or str(value).strip() == "" or str(value).strip() == "-":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    
    val_str = str(value).strip()
    # 1. 移除千分位点
    val_str = val_str.replace('.', '')
    # 2. 处理可能存在的逗号小数点
    val_str = val_str.replace(',', '.')
    # 3. 移除非数字字符 (保留负号和处理后的点)
    val_str = re.sub(r'[^0-9.\-]', '', val_str)
    
    try:
        return float(val_str) if val_str else 0.0
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
    if st.button("🚀 开始自动化核算"):
        try:
            # 读取数据
            def load_df(file):
                try: return pd.read_excel(file, dtype=str)
                except: return pd.read_csv(file, dtype=str)

            df_temp = load_df(f_a)
            df_order = load_df(f_b)
            df_income = load_df(f_c)
            df_cost = load_df(f_d)

            # --- 1. 定位关键列 ---
            oid_order = get_col_exact(df_order, 'order number')
            oid_income = get_col_exact(df_income, 'Order number')
            sku_order = get_col_exact(df_order, 'Nomor Referensi SKU')
            sku_cost = get_col_exact(df_cost, 'Nomor Referensi SKU')
            cost_val_col = get_col_exact(df_cost, '成本单价')
            
            # 订单表关键列
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')
            status_col = get_col_exact(df_order, 'Status Pesanan')

            if not oid_order or not oid_income:
                st.error("❌ 无法匹配订单号列，请检查表头。")
                st.stop()

            # --- 2. 局部金额清洗 (只针对确定是数字的列) ---
            # 清洗订单表
            for col in [price_col, voucher_col]:
                if col: df_order[col] = df_order[col].apply(clean_currency)
            df_order[qty_col] = pd.to_numeric(df_order[qty_col], errors='coerce').fillna(0)
            
            # 清洗收入表费用列
            income_fee_names = [
                'AMS Commission Fee', 
                'Commission fee (including PPN 10%)', 
                'Service Fee', 
                'Seller Order Processing Fee', 
                'Premium'
            ]
            income_fee_cols = []
            for name in income_fee_names:
                real_col = get_col_exact(df_income, name)
                if real_col:
                    df_income[real_col] = df_income[real_col].apply(clean_currency)
                    income_fee_cols.append(real_col)

            # 清洗成本表
            if cost_val_col:
                df_cost[cost_val_col] = df_cost[cost_val_col].apply(clean_currency)

            # --- 3. 关联与分摊计算 ---
            # 强制订单号为字符串
            df_order[oid_order] = df_order[oid_order].astype(str).str.strip()
            df_income[oid_income] = df_income[oid_income].astype(str).str.strip()
            
            # 计算每笔订单的商品行数
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            
            # 合并收入表 (只取订单号和费用列)
            df_income_sub = df_income[[oid_income] + income_fee_cols].drop_duplicates(oid_income)
            df_main = pd.merge(df_order, df_income_sub, left_on=oid_order, right_on=oid_income, how='left')
            
            # 合并成本表
            if sku_order and sku_cost and cost_val_col:
                df_cost_sub = df_cost[[sku_cost, cost_val_col]].drop_duplicates(sku_cost)
                df_main = pd.merge(df_main, df_cost_sub, left_on=sku_order, right_on=sku_cost, how='left')

            # --- 4. 财务公式应用 ---
            df_main = df_main.fillna(0)

            def calc_row(row):
                st_str = str(row.get(status_col, '')).lower()
                if 'batal' in st_str or 'cancel' in st_str:
                    return pd.Series([0.0] * 9)
                
                cnt = row['计数'] if row['计数'] > 0 else 1
                
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / cnt
                
                # 费用平摊计算
                fees_detail = []
                total_fees = 0
                for f_name in income_fee_names:
                    f_col = get_col_exact(df_income, f_name)
                    val = (row.get(f_col, 0) / cnt) if f_col else 0
                    fees_detail.append(val)
                    total_fees += val
                
                inc = s_amt - v_share + total_fees
                c_total = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return pd.Series([s_amt, v_share, inc, c_total] + fees_detail)

            res_cols = ['_S', '_V', '_I', '_C', '_f1', '_f2', '_f3', '_f4', '_f5']
            df_main[res_cols] = df_main.apply(calc_row, axis=1)

            # --- 5. 填充回模板 A ---
            df_final = pd.DataFrame(columns=df_temp.columns)
            special_map = {
                '成功订单销售金额': '_S',
                '优惠券': '_V',
                'income': '_I',
                '成本': cost_val_col,
                '总成本': '_C',
                '计数': '计数'
            }
            # 费用明细映射
            fee_map = {income_fee_names[i]: res_cols[4+i] for i in range(5)}

            for col in df_final.columns:
                c_name = str(col).strip()
                # 费用项平摊值
                found_fee = False
                for f_orig_name, f_calc_col in fee_map.items():
                    if f_orig_name.lower() in c_name.lower():
                        df_final[col] = df_main[f_calc_col]
                        found_fee = True
                        break
                if found_fee: continue
                
                # 计算项
                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    # 原始项 (再次清洗以防万一)
                    orig = get_col_exact(df_main, c_name)
                    if orig:
                        # 如果原始列也是金额类，确保它是干净的数字
                        if any(k in c_name.lower() for k in ['harga', 'diskon', 'voucher']):
                            df_final[col] = df_main[orig].apply(clean_currency)
                        else:
                            df_final[col] = df_main[orig]

            st.success("✅ 核算完成！")
            st.dataframe(df_final.head(10))

            output = io.BytesIO()
            df_final.to_excel(output, index=False)
            st.download_button("📥 下载结果报表", output.getvalue(), "Shopee_Final_A.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
            st.info("建议：请检查上传的表格是否有空行或损坏的单元格。")
