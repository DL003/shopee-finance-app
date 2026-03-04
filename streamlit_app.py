import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务四表合一", layout="wide")
st.title("📂 Shopee 财务账单自动化助手 (底表A生成模式)")

# --- 核心辅助：智能寻找列名 ---
def find_col(df, keyword):
    """在列名中智能寻找包含关键字的列"""
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            return col
    return None

# 1. 文件上传区
col1, col2 = st.columns(2)
with col1:
    f_b = st.file_uploader("1. 上传【表B：销售订单表】", type=['xlsx'])
    f_c = st.file_uploader("2. 上传【表C：订单收入报表】", type=['xlsx'])
with col2:
    f_d = st.file_uploader("3. 上传【表D：成本表】", type=['xlsx'])

if f_b and f_c and f_d:
    if st.button("🚀 开始按照表A逻辑生成汇总"):
        try:
            df_b = pd.read_excel(f_b)
            df_i = pd.read_excel(f_c) # 收入表
            df_cost = pd.read_excel(f_d) # 成本表

            # 步骤 1: 提取表 B 原始字段
            b_target = ['order number', 'Status Pesanan', 'Alasan Pembatalan', 'No. Resi', 
                        'Nomor Referensi SKU', 'Harga Setelah Diskon', 'Jumlah', 
                        'Returned quantity', 'Diskon Dari Shopee', 'Voucher Ditanggung Penjual']
            df_a = df_b[[c for c in b_target if c in df_b.columns]].copy()

            # 步骤 2: 计算计数
            df_a['计数'] = df_a.groupby('order number')['order number'].transform('count')

            # 步骤 3: 匹配表 C 费用 (模糊匹配关键字)
            df_i_clean = df_i[['order number']].copy()
            df_i_clean['fee_total'] = 0
            fee_keys = ['AMS Commission', 'Commission fee', 'Service Fee', 'Processing Fee', 'Premium']
            for k in fee_keys:
                found = find_col(df_i, k)
                if found:
                    df_i_clean['fee_total'] += df_i[found].fillna(0)
            
            df_a = pd.merge(df_a, df_i_clean.drop_duplicates('order number'), on='order number', how='left')

            # 步骤 4: 匹配表 D 成本
            sku_col = find_col(df_cost, "Nomor Referensi SKU")
            cost_col = find_col(df_cost, "成本")
            if sku_col and cost_col:
                df_a = pd.merge(df_a, df_cost[[sku_col, cost_col]], left_on='Nomor Referensi SKU', right_on=sku_col, how='left')
            else:
                st.warning("⚠️ 成本表未找到匹配列，请确保包含 'Nomor Referensi SKU' 和 '成本'。")

            # 步骤 5: 核心财务计算
            df_a = df_a.fillna(0)
            def run_calc(row):
                if str(row.get('Status Pesanan', '')).strip() == 'Batal':
                    return 0, 0, 0, 0
                s_amt = row.get('Harga Setelah Diskon', 0) * row.get('Jumlah', 0)
                v_share = row.get('Voucher Ditanggung Penjual', 0) / row['计数'] if row['计数'] > 0 else 0
                final_inc = s_amt - v_share + row.get('fee_total', 0)
                total_cost = row.get(cost_col, 0) * row.get('Jumlah', 0)
                return s_amt, v_share, final_inc, total_cost

            df_a[['成功订单销售金额', '优惠券', 'income', '最终成本']] = df_a.apply(lambda x: pd.Series(run_calc(x)), axis=1)
            df_a['ad'] = ""

            st.success("✅ 数据核算完成！")
            st.dataframe(df_a.head(20))

            output = io.BytesIO()
            df_a.to_excel(output, index=False)
            st.download_button("📥 下载汇总结果 Excel", output.getvalue(), "Shopee_Accounting_Final.xlsx")

        except Exception as e:
            st.error(f"处理报错: {e}")
