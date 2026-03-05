import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表核算", layout="wide")
st.title("📊 Shopee 财务账单自动化助手 (模板填充模式)")

# 辅助函数：模糊匹配列名
def find_col(df, keyword):
    for col in df.columns:
        if keyword.lower() in str(col).lower(): return col
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

            # --- 步骤 1: 以表B为核心准备基础数据 ---
            # 计算计数
            df_sales['计数'] = df_sales.groupby('order number')['order number'].transform('count')

            # --- 步骤 2: 匹配收入表C费用 ---
            fee_keys = ['AMS Commission', 'Commission fee', 'Service Fee', 'Processing Fee', 'Premium']
            df_i_clean = df_income[['order number']].copy()
            df_i_clean['fee_sum'] = 0
            for k in fee_keys:
                found = find_col(df_income, k)
                if found: df_i_clean['fee_sum'] += df_income[found].fillna(0)
            
            # 将收入数据合并到销售表
            df_main = pd.merge(df_sales, df_i_clean.drop_duplicates('order number'), on='order number', how='left')

            # --- 步骤 3: 匹配成本表D ---
            cost_sku = find_col(df_cost, "Nomor Referensi SKU")
            cost_val = find_col(df_cost, "成本")
            if cost_sku and cost_val:
                df_main = pd.merge(df_main, df_cost[[cost_sku, cost_val]], left_on='Nomor Referensi SKU', right_on=cost_sku, how='left')
            else:
                df_main['成本_unit'] = 0

            # --- 步骤 4: 执行财务计算 ---
            df_main = df_main.fillna(0)
            def run_calc(row):
                if str(row.get('Status Pesanan', '')).strip() == 'Batal':
                    return 0, 0, 0, 0
                s_amt = row.get('Harga Setelah Diskon', 0) * row.get('Jumlah', 0)
                v_share = row.get('Voucher Ditanggung Penjual', 0) / row['计数'] if row['计数'] > 0 else 0
                final_inc = s_amt - v_share + row.get('fee_sum', 0)
                total_cost = row.get(cost_val, 0) * row.get('Jumlah', 0)
                return s_amt, v_share, final_inc, total_cost

            df_main[['成功订单销售金额', '优惠券', 'income', '最终成本']] = df_main.apply(lambda x: pd.Series(run_calc(x)), axis=1)

            # --- 步骤 5: 将结果填充回模板 A 的结构 ---
            # 这一步会自动匹配模板 A 中存在的列名
            final_output = pd.DataFrame(columns=df_template.columns)
            for col in final_output.columns:
                if col in df_main.columns:
                    final_output[col] = df_main[col]
            
            # 如果模板里有 'ad' 留空
            if 'ad' in final_output.columns:
                final_output['ad'] = ""

            st.success("✅ 模板填充核算完成！")
            st.dataframe(final_output.head(20))

            # 导出
            out = io.BytesIO()
            final_output.to_excel(out, index=False)
            st.download_button("📥 下载填充后的表A", out.getvalue(), "Shopee_Final_TableA.xlsx")

        except Exception as e:
            st.error(f"逻辑执行报错: {e}")
