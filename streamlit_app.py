import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表核算", layout="wide")
st.title("📊 Shopee 财务账单自动化助手 (模糊匹配增强版)")

# --- 核心辅助：智能寻找列名 ---
def smart_find_col(df, keyword_list):
    """
    支持模糊关键字包含匹配。
    只要实际列名包含 keyword_list 中的任一关键字，即认为匹配成功。
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
            # 读取数据
            df_template = pd.read_excel(f_a)
            df_sales = pd.read_excel(f_b)
            df_income = pd.read_excel(f_c)
            df_cost = pd.read_excel(f_d)

            # 寻找订单号列
            sales_order_col = smart_find_col(df_sales, ['order number', '订单号'])
            income_order_col = smart_find_col(df_income, ['order number', '订单号'])
            
            if not sales_order_col or not income_order_col:
                st.error("❌ 无法识别订单号列，请确保销售表和收入表包含 'order number'。")
                st.stop()

            # --- 1. 定义费用关键字映射 ---
            # 这些关键字是表C中常见的核心文本，程序会忽略括号里的百分比进行匹配
            fee_mapping = {
                'AMS Commission Fee': ['AMS Commission Fee'],
                'Commission fee': ['Commission fee'],
                'Service Fee': ['Service Fee'],
                'Seller Order Processing Fee': ['Seller Order Processing Fee', 'Processing Fee'],
                'Premium': ['Premium']
            }

            # --- 2. 预处理收入表 C (提取费用并去重) ---
            df_income_processed = df_income[[income_order_col]].copy()
            # 记录在表C中实际找到的费用列名
            matched_fee_cols = {} 

            for internal_key, search_keys in fee_mapping.items():
                real_col = smart_find_col(df_income, search_keys)
                if real_col:
                    df_income_processed[internal_key] = pd.to_numeric(df_income[real_col], errors='coerce').fillna(0)
                    matched_fee_cols[internal_key] = internal_key
            
            # 按订单号去重，确保每个订单只有一行费用汇总
            df_income_processed = df_income_processed.drop_duplicates(income_order_col)

            # --- 3. 准备成本表 D 数据 ---
            cost_sku_col = smart_find_col(df_cost, ['Nomor Referensi SKU', 'SKU'])
            cost_val_col = smart_find_col(df_cost, ['成本', '单价'])
            sales_sku_col = smart_find_col(df_sales, ['Nomor Referensi SKU', 'SKU'])

            # --- 4. 核心合并与计算 ---
            # A. 计算表B每订单行数（平摊优惠券用）
            df_sales['计数'] = df_sales.groupby(sales_order_col)[sales_order_col].transform('count')
            
            # B. 合并收入数据
            df_main = pd.merge(df_sales, df_income_processed, left_on=sales_order_col, right_on=income_order_col, how='left')
            
            # C. 合并成本数据
            if cost_sku_col and cost_val_col and sales_sku_col:
                df_main = pd.merge(df_main, df_cost[[cost_sku_col, cost_val_col]], left_on=sales_sku_col, right_on=cost_sku_col, how='left')
            
            # D. 执行具体财务公式
            status_col = smart_find_col(df_sales, ['Status Pesanan', '订单状态'])
            price_col = smart_find_col(df_sales, ['Harga Setelah Diskon', '折后价'])
            qty_col = smart_find_col(df_sales, ['Jumlah', '数量'])
            voucher_col = smart_find_col(df_sales, ['Voucher Ditanggung Penjual', '优惠券'])

            df_main = df_main.fillna(0)
            
            def run_calc(row):
                # 判定取消状态
                status_str = str(row.get(status_col, '')).lower()
                if any(x in status_str for x in ['batal', 'cancelled', '取消']):
                    return 0, 0, 0, 0
                
                # 计算逻辑
                success_sales = row.get(price_col, 0) * row.get(qty_col, 0)
                voucher_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                
                # 汇总所有匹配到的费用（计算 income 用）
                total_fees = sum([row.get(k, 0) for k in fee_mapping.keys()])
                
                final_income = success_sales - voucher_share + total_fees
                item_cost = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return success_sales, voucher_share, final_income, item_cost

            df_main[['_calc_sales', '_calc_voucher', '_calc_income', '_calc_cost']] = df_main.apply(
                lambda x: pd.Series(run_calc(x)), axis=1
            )

            # --- 5. 最终匹配并填充回模板 A ---
            final_df = pd.DataFrame(columns=df_template.columns)
            
            for col in final_df.columns:
                col_name = str(col).strip()
                
                # 情况 A: 匹配费用明细列 (模糊识别文本)
                matched_fee = None
                for k in fee_mapping.keys():
                    if k.lower() in col_name.lower():
                        matched_fee = k
                        break
                
                if matched_fee:
                    final_df[col] = df_main[matched_fee]
                # 情况 B: 匹配计算结果列
                elif '成功订单销售金额' in col_name: final_df[col] = df_main['_calc_sales']
                elif 'income' in col_name.lower(): final_df[col] = df_main['_calc_income']
                elif '最终成本' in col_name or ('成本' in col_name and '单价' not in col_name): 
                    final_df[col] = df_main['_calc_cost']
                elif '优惠券' in col_name: final_df[col] = df_main['_calc_voucher']
                # 情况 C: 匹配其他原始列
                else:
                    orig_match = smart_find_col(df_main, [col_name])
                    if orig_match: final_df[col] = df_main[orig_match]

            st.success("✅ 核算成功！费用列已根据文字关键字自动匹配。")
            st.dataframe(final_df.head(10))

            # 导出 Excel
            output = io.BytesIO()
            final_df.to_excel(output, index=False)
            st.download_button("📥 下载核算结果 (表A)", output.getvalue(), "Shopee_Accounting_Result.xlsx")

        except Exception as e:
            st.error(f"处理中发生逻辑错误: {e}")
