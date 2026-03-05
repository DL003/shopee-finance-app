import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表核算", layout="wide")
st.title("📊 Shopee 财务账单自动化助手 (增强匹配版)")

# --- 核心辅助：超级匹配函数 ---
def smart_find_col(df, keyword_list):
    """
    更加智能的列名寻找：
    1. 尝试完全匹配（去除空格）
    2. 尝试关键字包含匹配
    """
    # 预处理所有列名：转小写并去空格
    clean_cols = {str(c).strip().lower(): c for c in df.columns}
    
    # 1. 尝试在清洗后的列名中找完全匹配
    for key in keyword_list:
        if key.lower() in clean_cols:
            return clean_cols[key.lower()]
            
    # 2. 尝试模糊包含匹配
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

            # --- 寻找关键列名 (解决报错的核心) ---
            # 找销售表里的订单号
            sales_order_col = smart_find_col(df_sales, ['order number', '订单号', 'No. Pesanan'])
            # 找收入表里的订单号
            income_order_col = smart_find_col(df_income, ['order number', '订单号', 'No. Pesanan'])
            
            if not sales_order_col or not income_order_col:
                st.error(f"❌ 找不到订单号列！销售表识别为: {sales_order_col}, 收入表识别为: {income_order_col}")
                st.stop()

            # --- 步骤 1: 准备基础数据 ---
            # 计算计数
            df_sales['计数'] = df_sales.groupby(sales_order_col)[sales_order_col].transform('count')

            # --- 步骤 2: 匹配收入表C费用 ---
            # 费用关键词
            fee_keys = ['AMS Commission', 'Commission fee', 'Service Fee', 'Processing Fee', 'Premium']
            df_income_clean = df_income[[income_order_col]].copy()
            df_income_clean['fee_sum'] = 0
            
            for k in fee_keywords:
                found = smart_find_col(df_income, [k])
                if found:
                    df_income_clean['fee_sum'] += pd.to_numeric(df_income[found], errors='coerce').fillna(0)
            
            # 将收入数据合并到销售表 (左连接)
            df_main = pd.merge(df_sales, df_income_clean.drop_duplicates(income_order_col), 
                               left_on=sales_order_col, right_on=income_order_col, how='left')

            # --- 步骤 3: 匹配成本表D ---
            cost_sku_col = smart_find_col(df_cost, ['Nomor Referensi SKU', 'SKU', '货号'])
            cost_val_col = smart_find_col(df_cost, ['成本', '单价', 'Cost'])
            
            sales_sku_col = smart_find_col(df_sales, ['Nomor Referensi SKU', 'SKU'])
            
            if cost_sku_col and cost_val_col and sales_sku_col:
                df_main = pd.merge(df_main, df_cost[[cost_sku_col, cost_val_col]], 
                                   left_on=sales_sku_col, right_on=cost_sku_col, how='left')
            else:
                df_main['成本_val'] = 0

            # --- 步骤 4: 执行财务计算 ---
            # 状态列寻找
            status_col = smart_find_col(df_sales, ['Status Pesanan', '订单状态'])
            price_col = smart_find_col(df_sales, ['Harga Setelah Diskon', '折后价'])
            qty_col = smart_find_col(df_sales, ['Jumlah', '数量'])
            voucher_col = smart_find_col(df_sales, ['Voucher Ditanggung Penjual', '优惠券'])

            df_main = df_main.fillna(0)
            
            def run_calc(row):
                # 判定取消状态
                status_str = str(row.get(status_col, '')).lower()
                if 'batal' in status_str or 'cancelled' in status_str or '取消' in status_str:
                    return 0, 0, 0, 0
                
                # 计算逻辑
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                v_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                final_inc = s_amt - v_share + row.get('fee_sum', 0)
                total_cost = row.get(cost_val_col, 0) * row.get(qty_col, 0)
                
                return s_amt, v_share, final_inc, total_cost

            df_main[['成功订单销售金额', '优惠券', 'income', '最终成本']] = df_main.apply(
                lambda x: pd.Series(run_calc(x)), axis=1
            )

            # --- 步骤 5: 填充回模板 A ---
            # 自动映射：将 df_main 中有的列，填入 df_template 对应的列中
            final_output = pd.DataFrame(columns=df_template.columns)
            for col in final_output.columns:
                # 尝试模糊匹配模板列名
                matched_main_col = smart_find_col(df_main, [col])
                if matched_main_col:
                    final_output[col] = df_main[matched_main_col]

            st.success("✅ 数据核算并模板填充完成！")
            st.dataframe(final_output.head(10))

            # 导出
            out = io.BytesIO()
            final_output.to_excel(out, index=False)
            st.download_button("📥 下载填充后的表A", out.getvalue(), "Shopee_Final_A.xlsx")

        except Exception as e:
            st.error(f"逻辑执行报错: {e}")
