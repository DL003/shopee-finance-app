import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表核算", layout="wide")
st.title("📊 Shopee 财务账单自动化助手 (模糊匹配增强版)")

# --- 核心辅助：智能寻找列名 ---
def smart_find_col(df, keyword_list):
    """
    支持模糊关键字包含匹配。只要实际列名包含 keyword_list 中的任一字符，即认为匹配。
    """
    for actual_col in df.columns:
        actual_col_str = str(actual_col).strip().lower()
        for key in keyword_list:
            if key.lower() in actual_col_str:
                return actual_col
    return None

# 文件上传区
col1, col2 = st.columns(2)
with col1:
    f_a = st.file_uploader("1. 上传【表A：底表模板】", type=['xlsx'])
    f_b = st.file_uploader("2. 上传【表B：销售订单表】", type=['xlsx'])
with col2:
    f_c = st.file_uploader("3. 上传【表C：订单收入报表】", type=['xlsx'])
    f_d = st.file_uploader("4. 上传【表D：成本表】", type=['xlsx'])

if f_a and f_b and f_c and f_d:
    if st.button("🚀 开始按照模板填充数据"):
        try:
            # 读取原始数据
            df_template = pd.read_excel(f_a)
            df_sales = pd.read_excel(f_b)
            df_income = pd.read_excel(f_c)
            df_cost = pd.read_excel(f_d)

            # 寻找订单号列
            sales_order_col = smart_find_col(df_sales, ['order number', '订单号'])
            income_order_col = smart_find_col(df_income, ['order number', '订单号'])
            
            if not sales_order_col or not income_order_col:
                st.error("❌ 无法识别订单号列，请确保销售表和收入表包含 'order number' 列。")
                st.stop()

            # --- 1. 定义费用关键字映射 ---
            # 这里的 Key 是模板 A 里的列名关键字，Value 是我们去收入表 C 里寻找的模糊关键字
            fee_mapping = {
                'AMS Commission': ['AMS Commission'],
                'Commission fee': ['Commission fee'],
                'Service Fee': ['Service Fee'],
                'Processing Fee': ['Processing Fee', 'Seller Order Processing'],
                'Premium': ['Premium']
            }

            # --- 2. 预处理收入表 C ---
            # 根据模糊匹配，在表 C 中找到对应的真实列名，并重命名，方便后续合并
            df_income_sub = df_income[[income_order_col]].copy()
            income_real_cols = {} # 存储 模板列名 -> 表C真实列名 的映射
            
            for template_key, search_keys in fee_mapping.items():
                real_col = smart_find_col(df_income, search_keys)
                if real_col:
                    # 将该列数据转换成数值，并存入清洗后的收入表
                    df_income_sub[template_key] = pd.to_numeric(df_income[real_col], errors='coerce').fillna(0)
                    income_real_cols[template_key] = template_key
            
            # 去重订单号，防止匹配出多行
            df_income_sub = df_income_sub.drop_duplicates(income_order_col)

            # --- 3. 匹配成本表 D ---
            cost_sku_col = smart_find_col(df_cost, ['Nomor Referensi SKU', 'SKU'])
            cost_val_col = smart_find_col(df_cost, ['成本', '单价'])
            sales_sku_col = smart_find_col(df_sales, ['Nomor Referensi SKU', 'SKU'])

            # --- 4. 核心合并逻辑 ---
            # 计算表 B 计数
            df_sales['计数'] = df_sales.groupby(sales_order_col)[sales_order_col].transform('count')
            
            # 合并收入数据
            df_main = pd.merge(df_sales, df_income_sub, left_on=sales_order_col, right_on=income_order_col, how='left')
            
            # 合并成本数据
            if cost_sku_col and cost_val_col and sales_sku_col:
                df_main = pd.merge(df_main, df_cost[[cost_sku_col, cost_val_col]], left_on=sales_sku_col, right_on=cost_sku_col, how='left')
            
            # --- 5. 财务计算逻辑 ---
            status_col = smart_find_col(df_sales, ['Status Pesanan', '订单状态'])
            price_col = smart_find_col(df_sales, ['Harga Setelah Diskon', '折后价'])
            qty_col = smart_find_col(df_sales, ['Jumlah', '数量'])
            voucher_col = smart_find_col(df_sales, ['Voucher Ditanggung Penjual', '优惠券'])

            df_main = df_main.fillna(0)
            
            def run_calc(row):
                status_str = str(row.get(status_col, '')).lower()
                if any(x in status_str for x in ['batal', 'cancelled', '取消']):
                    return 0, 0, 0, 0
                
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                
                # 计算总费用（用于计算 income）
                current_fee_sum = sum([row.get(k, 0) for k in fee_mapping.keys()])
                
                final_inc = s_amt - v_share + current_fee_sum
                total_cost = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                return s_amt, v_share, final_inc, total_cost

            df_main[['calc_sales', 'calc_voucher', 'calc_income', 'calc_cost']] = df_main.apply(
                lambda x: pd.Series(run_calc(x)), axis=1
            )

            # --- 6. 填充回模板 A ---
            final_output = pd.DataFrame(columns=df_template.columns)
            for col in final_output.columns:
                col_str = str(col).strip()
                # A. 先看是不是我们定义的费用列（模糊匹配文字）
                matched_fee_key = None
                for k in fee_mapping.keys():
                    if k.lower() in col_str.lower():
                        matched_fee_key = k
                        break
                
                if matched_fee_key:
                    final_output[col] = df_main[matched_fee_key]
                # B. 再看是不是我们计算出的核心结果列
                elif '成功订单销售金额' in col_str: final_output[col] = df_main['calc_sales']
                elif 'income' in col_str.lower(): final_output[col] = df_main['calc_income']
                elif '成本' in col_str and '最终' not in col_str: final_output[col] = df_main['calc_cost']
                elif '优惠券' in col_str: final_output[col] = df_main['calc_voucher']
                # C. 最后尝试直接匹配原始列名
                else:
                    orig_col = smart_find_col(df_main, [col])
                    if orig_col: final_output[col] = df_main[orig_col]

            st.success("✅ 核算完成！已启用模糊表头匹配逻辑。")
            st.dataframe(final_output.head(10))

            out = io.BytesIO()
            final_output.to_excel(out, index=False)
            st.download_button("📥 下载填充后的表A", out.getvalue(), "Shopee_Accounting_Final.xlsx")

        except Exception as e:
            st.error(f"逻辑错误: {e}")
