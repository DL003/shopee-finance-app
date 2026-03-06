import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表核算", layout="wide")
st.title("📊 Shopee 财务账单自动化助手 (最终修正版)")

# --- 核心辅助：超级匹配函数 ---
def smart_find_col(df, keyword_list):
    """
    智能列名寻找：支持去空格完全匹配和模糊包含匹配
    """
    clean_cols = {str(c).strip().lower(): c for c in df.columns}
    # 1. 完全匹配
    for key in keyword_list:
        if key.lower() in clean_cols:
            return clean_cols[key.lower()]
    # 2. 模糊匹配
    for key in keyword_list:
        for actual_col in df.columns:
            if key.lower() in str(actual_col).lower():
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

            # --- 1. 寻找关键列名 ---
            sales_order_col = smart_find_col(df_sales, ['order number', '订单号', 'No. Pesanan'])
            income_order_col = smart_find_col(df_income, ['order number', '订单号', 'No. Pesanan'])
            
            if not sales_order_col or not income_order_col:
                st.error(f"❌ 找不到订单号列！请检查销售表和收入表。")
                st.stop()

            # --- 2. 准备基础数据 ---
            # 计算计数
            df_sales['计数'] = df_sales.groupby(sales_order_col)[sales_order_col].transform('count')

            # --- 3. 匹配收入表C费用 (修正变量名错误并加强匹配) ---
            target_fee_keys = ['AMS Commission', 'Commission fee', 'Service Fee', 'Processing Fee', 'Premium']
            df_income_clean = df_income[[income_order_col]].copy()
            df_income_clean['fee_sum'] = 0
            
            # 遍历每个关键词，累加匹配到的列
            for k in target_fee_keys:
                found_col = smart_find_col(df_income, [k])
                if found_col:
                    df_income_clean['fee_sum'] += pd.to_numeric(df_income[found_col], errors='coerce').fillna(0)
            
            # 合并收入数据到销售主表
            df_main = pd.merge(df_sales, df_income_clean.drop_duplicates(income_order_col), 
                               left_on=sales_order_col, right_on=income_order_col, how='left')

            # --- 4. 匹配成本表D ---
            cost_sku_col = smart_find_col(df_cost, ['Nomor Referensi SKU', 'SKU', '货号'])
            cost_val_col = smart_find_col(df_cost, ['成本', '单价', 'Cost'])
            sales_sku_col = smart_find_col(df_sales, ['Nomor Referensi SKU', 'SKU'])
            
            if cost_sku_col and cost_val_col and sales_sku_col:
                df_main = pd.merge(df_main, df_cost[[cost_sku_col, cost_val_col]], 
                                   left_on=sales_sku_col, right_on=cost_sku_col, how='left')
            else:
                st.warning("⚠️ 成本表或销售表的 SKU/成本 列未完全匹配，可能导致成本为 0。")

            # --- 5. 执行财务计算 ---
            status_col = smart_find_col(df_sales, ['Status Pesanan', '订单状态'])
            price_col = smart_find_col(df_sales, ['Harga Setelah Diskon', '折后价'])
            qty_col = smart_find_col(df_sales, ['Jumlah', '数量'])
            voucher_col = smart_find_col(df_sales, ['Voucher Ditanggung Penjual', '优惠券'])

            df_main = df_main.fillna(0)
            
            def run_calc(row):
                # 判定取消状态 (支持多语种关键字)
                status_str = str(row.get(status_col, '')).lower()
                if any(x in status_str for x in ['batal', 'cancelled', '取消']):
                    return 0, 0, 0, 0
                
                # 计算逻辑
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                # Income = 销售额 - 优惠券分摊 + 平台费总和 (平台费通常为负数)
                final_inc = s_amt - v_share + row.get('fee_sum', 0)
                # 成本计算 = SKU单价 * 数量
                total_cost = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return s_amt, v_share, final_inc, total_cost

            df_main[['成功订单销售金额', '优惠券', 'income', '成本']] = df_main.apply(
                lambda x: pd.Series(run_calc(x)), axis=1
            )

            # --- 6. 填充回模板 A ---
            final_output = pd.DataFrame(columns=df_template.columns)
            for col in final_output.columns:
                # 智能寻找主表中对应的计算结果或原始数据
                matched_main_col = smart_find_col(df_main, [col])
                if matched_main_col:
                    final_output[col] = df_main[matched_main_col]

            st.success("✅ 核算底表生成成功！")
            st.dataframe(final_output.head(10))

            # 导出
            out = io.BytesIO()
            final_output.to_excel(out, index=False)
            st.download_button("📥 下载核算结果 (Excel)", out.getvalue(), "Shopee_Final_A.xlsx")

        except Exception as e:
            st.error(f"发生未预料的逻辑错误: {e}")
