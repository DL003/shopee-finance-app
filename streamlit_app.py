import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务核算工具", layout="wide")

st.title("📊 Shopee 财务账单自动核算 (优化版)")

# --- 辅助函数：模糊匹配列名 ---
def find_col(df, keyword):
    """在列名中寻找包含关键字的列名"""
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            return col
    return None

# 文件上传
col1, col2, col3 = st.columns(3)
with col1:
    f_sales = st.file_uploader("1. 销售订单表", type=['xlsx'])
with col2:
    f_income = st.file_uploader("2. 订单收入报表", type=['xlsx'])
with col3:
    f_cost = st.file_uploader("3. 成本表", type=['xlsx'])

if f_sales and f_income and f_cost:
    if st.button("🚀 开始自动化对账"):
        try:
            df_s = pd.read_excel(f_sales)
            df_i = pd.read_excel(f_income)
            df_c = pd.read_excel(f_cost)

            # 清理列名空格
            df_s.columns = df_s.columns.str.strip()
            df_i.columns = df_i.columns.str.strip()
            df_c.columns = df_c.columns.str.strip()

            # 动态寻找收入表中的费用列
            # 我们通过关键字匹配，防止因为比例变化（如7.2%变8%）导致报错
            col_order = find_col(df_i, "order number")
            col_ams = find_col(df_i, "AMS Commission")
            col_comm = find_col(df_i, "Commission fee")
            col_serv = find_col(df_i, "Service Fee")
            col_proc = find_col(df_i, "Processing Fee")
            col_prem = find_col(df_i, "Premium")

            if not col_order:
                st.error("❌ 在收入报表中没找到 'order number' 列，请检查文件。")
                st.stop()

            # 提取并重命名，方便计算
            df_i_clean = df_i[[col_order]].copy()
            df_i_clean['fee_total'] = 0
            
            for c in [col_ams, col_comm, col_serv, col_proc, col_prem]:
                if c:
                    df_i_clean['fee_total'] += df_i[c].fillna(0)

            # 匹配逻辑
            df_s['计数'] = df_s.groupby('order number')['order number'].transform('count')
            df_m = pd.merge(df_s, df_i_clean.drop_duplicates(col_order), 
                            left_on='order number', right_on=col_order, how='left')
            
            # 成本匹配
            # 这里也用了模糊匹配，防止“成本单价”写错
            cost_sku_col = find_col(df_c, "Nomor Referensi SKU")
            cost_price_col = find_col(df_c, "成本") 
            
            if cost_sku_col and cost_price_col:
                df_m = pd.merge(df_m, df_c[[cost_sku_col, cost_price_col]], 
                                left_on='Nomor Referensi SKU', right_on=cost_sku_col, how='left')
            else:
                st.warning("⚠️ 成本表匹配失败：请检查列名是否包含 'Nomor Referensi SKU' 和 '成本'")

            # 计算
            def run_calc(row):
                if row['Status Pesanan'] == 'Batal':
                    return 0, 0, 0, 0
                s_amt = row.get('Harga Setelah Diskon', 0) * row.get('Jumlah', 0)
                v_share = row.get('Voucher Ditanggung Penjual', 0) / row['计数'] if row['计数'] > 0 else 0
                # 最终收入 = 销售额 - 优惠券分摊 + 平台费总和
                final_inc = s_amt - v_share + row.get('fee_total', 0)
                c_total = row.get(cost_price_col, 0) * row.get('Jumlah', 0) if pd.notnull(row.get(cost_price_col)) else 0
                return s_amt, v_share, final_inc, c_total

            df_m[['成功订单销售金额', '优惠券', 'income', '成本']] = df_m.apply(lambda x: pd.Series(run_calc(x)), axis=1)
            df_m['ad'] = ""

            st.success("✅ 匹配计算成功！")
            st.dataframe(df_m.head(10))

            output = io.BytesIO()
            df_m.to_excel(output, index=False)
            st.download_button("📥 下载结果 Excel", output.getvalue(), "Result_Optimized.xlsx")

        except Exception as e:
            st.error(f"发生错误: {e}")
