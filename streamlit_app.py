import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (广告平摊+刷单处理版)")

# --- 核心辅助：清洗金额格式 ---
def clean_currency(value):
    if pd.isna(value) or str(value).strip() == "" or str(value).strip() == "-":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    val_str = str(value).strip().replace('.', '').replace(',', '.')
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

# --- UI 布局 ---
col1, col2 = st.columns(2)
with col1:
    f_a = st.file_uploader("1. 上传【模板表.xlsx】", type=['xlsx', 'csv'])
    f_b = st.file_uploader("2. 上传【订单表.xlsx】", type=['xlsx', 'csv'])
    total_ad_input = st.number_input("💰 请输入本月总广告费 (用于按销售额平摊):", min_value=0.0, value=0.0, step=100.0)
with col2:
    f_c = st.file_uploader("3. 上传【订单收入.xlsx】", type=['xlsx', 'csv'])
    f_d = st.file_uploader("4. 上传【成本表.xlsx】", type=['xlsx', 'csv'])
    f_e = st.file_uploader("5. 上传【表E：刷单表】(选填)", type=['xlsx', 'csv'])

if f_a and f_b and f_c and f_d:
    if st.button("🚀 开始执行全自动化核算"):
        try:
            def load_df(file):
                try: return pd.read_excel(file, dtype=str)
                except: return pd.read_csv(file, dtype=str)

            df_temp = load_df(f_a)
            df_order = load_df(f_b)
            df_income = load_df(f_c)
            df_cost = load_df(f_d)
            # 处理刷单表
            df_brush = load_df(f_e) if f_e else pd.DataFrame(columns=['order number'])
            brush_list = set(df_brush[get_col_exact(df_brush, 'order number')].astype(str).str.strip().tolist()) if not df_brush.empty else set()

            # --- 1. 定位关键列 ---
            oid_order = get_col_exact(df_order, 'order number')
            oid_income = get_col_exact(df_income, 'Order number')
            sku_order = get_col_exact(df_order, 'Nomor Referensi SKU')
            sku_cost = get_col_exact(df_cost, 'Nomor Referensi SKU')
            cost_val_col = get_col_exact(df_cost, '成本单价')
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')
            status_col = get_col_exact(df_order, 'Status Pesanan')

            # --- 2. 金额清洗 ---
            for col in [price_col, voucher_col]:
                df_order[col] = df_order[col].apply(clean_currency)
            df_order[qty_col] = pd.to_numeric(df_order[qty_col], errors='coerce').fillna(0)
            
            fee_names = ['AMS Commission Fee', 'Commission fee (including PPN 10%)', 'Service Fee', 'Seller Order Processing Fee', 'Premium']
            fee_cols_found = []
            for name in fee_names:
                real_c = get_col_exact(df_income, name)
                if real_c:
                    df_income[real_c] = df_income[real_c].apply(clean_currency)
                    fee_cols_found.append(real_c)

            df_cost[cost_val_col] = df_cost[cost_val_col].apply(clean_currency)

            # --- 3. 合并数据 ---
            df_order[oid_order] = df_order[oid_order].astype(str).str.strip()
            df_income[oid_income] = df_income[oid_income].astype(str).str.strip()
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            
            df_income_sub = df_income[[oid_income] + fee_cols_found].drop_duplicates(oid_income)
            df_main = pd.merge(df_order, df_income_sub, left_on=oid_order, right_on=oid_income, how='left')
            
            if sku_order and sku_cost:
                df_cost_sub = df_cost[[sku_cost, cost_val_col]].drop_duplicates(sku_cost)
                df_main = pd.merge(df_main, df_cost_sub, left_on=sku_order, right_on=sku_cost, how='left')

            df_main = df_main.fillna(0)

            # --- 4. 核心计算 (包含刷单逻辑) ---
            def calc_logic(row):
                oid = str(row.get(oid_order, ''))
                st_str = str(row.get(status_col, '')).lower()
                is_batal = 'batal' in st_str or 'cancel' in st_str
                is_brush = oid in brush_list
                
                cnt = row['计数'] if row['计数'] > 0 else 1
                
                # 计算费用分摊 (无论是否刷单都要计入)
                f_sum = 0
                f_list = []
                for f_name in fee_names:
                    real_c = get_col_exact(df_income, f_name)
                    val = (row.get(real_c, 0) / cnt) if real_c else 0
                    f_list.append(val)
                    f_sum += val

                if is_batal:
                    return pd.Series([0.0, 0.0, 0.0, 0.0] + [0.0]*5)
                
                # 初始计算项
                raw_s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / cnt
                c_total = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                if is_brush:
                    # 刷单逻辑：销售额=0, 成本=0, 收入=费用总和
                    return pd.Series([0.0, 0.0, f_sum, 0.0] + f_list)
                else:
                    # 正常逻辑
                    inc = raw_s_amt - v_share + f_sum
                    return pd.Series([raw_s_amt, v_share, inc, c_total] + f_list)

            res_cols = ['_S', '_V', '_I', '_C', '_f1', '_f2', '_f3', '_f4', '_f5']
            df_main[res_cols] = df_main.apply(calc_logic, axis=1)

            # --- 5. 广告费平摊逻辑 ---
            total_sales_volume = df_main['_S'].sum()
            if total_sales_volume > 0:
                df_main['_AD'] = (df_main['_S'] / total_sales_volume) * total_ad_input
            else:
                df_main['_AD'] = 0.0

            # --- 6. 填充模板 ---
            df_final = pd.DataFrame(columns=df_temp.columns)
            special_map = {'成功订单销售金额': '_S', '优惠券': '_V', 'income': '_I', '成本': cost_val_col, '总成本': '_C', 'ad': '_AD'}
            fee_map = {fee_names[i]: res_cols[4+i] for i in range(5)}

            for col in df_final.columns:
                c_name = str(col).strip()
                # 费用列
                found_fee = False
                for f_orig, f_calc in fee_map.items():
                    if f_orig.lower() in c_name.lower():
                        df_final[col] = df_main[f_calc]
                        found_fee = True; break
                if found_fee: continue

                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    orig = get_col_exact(df_main, c_name)
                    if orig: df_final[col] = df_main[orig]

            st.success(f"✅ 核算完成！广告费已按销售额比例分摊至各SKU。")
            if not brush_list: st.warning("⚠️ 未上传刷单表或未识别到刷单订单，按常规订单处理。")
            st.dataframe(df_final.head(10))

            out = io.BytesIO()
            df_final.to_excel(out, index=False)
            st.download_button("📥 下载最终对账报表 (带广告+刷单处理)", out.getvalue(), "Shopee_Final_Full_Report.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
