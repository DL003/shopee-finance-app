import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (费用平摊+格式修复版)")

# --- 核心辅助：清洗印尼/特殊格式金额 ---
def clean_currency(value):
    """
    深度清洗金额：将 '125.795' 还原为 125795.0
    逻辑：强制移除所有点（千位符），处理逗号
    """
    if pd.isna(value) or str(value).strip() == "" or str(value).strip() == "-":
        return 0.0
    
    val_str = str(value).strip()
    
    # 1. 移除千分位点
    val_str = val_str.replace('.', '')
    # 2. 处理可能存在的逗号小数点
    val_str = val_str.replace(',', '.')
    # 3. 移除非数字字符 (保留负号和处理后的点)
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
    # 关键词模糊匹配
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
    if st.button("🚀 开始核算 (包含费用平摊修复)"):
        try:
            # 读取数据：强制字符串读取以防丢失千分位点
            def load_df(file):
                try: return pd.read_excel(file, dtype=str)
                except: return pd.read_csv(file, dtype=str)

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

            # 1. 预处理收入表费用
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
                    df_income_clean[key] = df_income[actual_col].apply(clean_currency)
                else:
                    df_income_clean[key] = 0.0
            
            df_income_final = df_income_clean.drop_duplicates(oid_income)

            # 2. 预处理订单表金额
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')
            
            df_order[price_col] = df_order[price_col].apply(clean_currency)
            df_order[voucher_col] = df_order[voucher_col].apply(clean_currency)
            df_order[qty_col] = pd.to_numeric(df_order[qty_col], errors='coerce').fillna(0)

            # 3. 关联数据
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            df_main = pd.merge(df_order, df_income_final, left_on=oid_order, right_on=oid_income, how='left')
            
            # 关联成本
            if sku_order and sku_cost and cost_val_col:
                df_cost_sub = df_cost[[sku_cost, cost_val_col]].copy()
                df_cost_sub[cost_val_col] = df_cost_sub[cost_val_col].apply(clean_currency)
                df_main = pd.merge(df_main, df_cost_sub, left_on=sku_order, right_on=sku_cost, how='left')
            else:
                df_main['成本单价'] = 0.0

            # 4. 核心核算逻辑 (包含分摊逻辑)
            status_col = get_col_exact(df_order, 'Status Pesanan')
            
            def calc_row(row):
                st_str = str(row.get(status_col, '')).lower()
                if 'batal' in st_str or 'cancel' in st_str:
                    return pd.Series([0.0]*9)
                
                cnt = row['计数'] if row['计数'] > 0 else 1
                
                # 销售额 = 折后价 * 数量
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                # 优惠券平摊
                v_share = row.get(voucher_col, 0) / cnt
                
                # --- 费用平摊处理 ---
                f_ams = row.get('ams', 0) / cnt
                f_comm = row.get('comm', 0) / cnt
                f_service = row.get('service', 0) / cnt
                f_proc = row.get('proc', 0) / cnt
                f_prem = row.get('premium', 0) / cnt
                
                # 计算 income = 销售额 - 优惠券分摊 + 平台费分摊 (注：费用通常为负数)
                inc = s_amt - v_share + (f_ams + f_comm + f_service + f_proc + f_prem)
                
                # 成本计算
                c_total = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return pd.Series([s_amt, v_share, inc, c_total, f_ams, f_comm, f_service, f_proc, f_prem])

            # 应用计算
            calc_res_cols = ['_S', '_V', '_I', '_C', '_f_ams', '_f_comm', '_f_service', '_f_proc', '_f_prem']
            df_main[calc_res_cols] = df_main.apply(calc_row, axis=1)

            # 5. 填充回模板 A
            df_final = pd.DataFrame(columns=df_temp.columns)
            
            # 费用明细结果映射
            fee_result_map = {
                'ams': '_f_ams',
                'comm': '_f_comm',
                'service': '_f_service',
                'proc': '_f_proc',
                'premium': '_f_prem'
            }
            
            # 特殊结果列映射
            special_map = {
                '成功订单销售金额': '_S',
                '优惠券': '_V',
                'income': '_I',
                '成本': cost_val_col,
                '总成本': '_C'
            }

            for col in df_final.columns:
                c_name = str(col).strip()
                
                # A. 费用明细列 (填入平摊后的数据)
                matched_fee = False
                for k_internal, k_orig in income_fee_map.items():
                    if k_orig.lower() in c_name.lower():
                        df_final[col] = df_main[fee_result_map[k_internal]]
                        matched_fee = True
                        break
                if matched_fee: continue

                # B. 特殊计算列
                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    # C. 原始字段直接填充
                    orig = get_col_exact(df_main, c_name)
                    if orig: df_final[col] = df_main[orig]

            st.success("✅ 核算完成！费用已根据商品行数自动平摊。")
            st.dataframe(df_final.head(10))

            output = io.BytesIO()
            df_final.to_excel(output, index=False)
            st.download_button("📥 下载平摊修复后的表A", output.getvalue(), "Shopee_Final_A_Allocated.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
