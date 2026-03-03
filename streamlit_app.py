import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务核算工具", layout="wide")

st.title("📊 Shopee 财务账单自动核算")
st.markdown("---")

# 侧边栏：操作说明
with st.sidebar:
    st.header("使用说明")
    st.write("1. 依次上传三个 Excel 文件")
    st.write("2. 点击开始处理按钮")
    st.write("3. 预览结果并下载 Excel")
    st.info("提示：系统会自动根据 'order number' 进行多表匹配，并根据订单状态处理数据。")

# 1. 文件上传
col1, col2, col3 = st.columns(3)
with col1:
    f_sales = st.file_uploader("上传【销售订单表】", type=['xlsx'])
with col2:
    f_income = st.file_uploader("上传【订单收入报表】", type=['xlsx'])
with col3:
    f_cost = st.file_uploader("上传【本地成本表】", type=['xlsx'])

# 2. 核心处理逻辑
if f_sales and f_income and f_cost:
    if st.button("🚀 开始自动化对账"):
        try:
            # 读取数据
            df_s = pd.read_excel(f_sales)
            df_i = pd.read_excel(f_income)
            df_c = pd.read_excel(f_cost)

            # 计算订单行计数（用于平摊优惠券）
            df_s['计数'] = df_s.groupby('order number')['order number'].transform('count')

            # 匹配收入表字段
            income_fields = ['order number', 'AMS Commission Fee', 'Commission fee (including PPN 10%)(7.2%)', 
                             'Service Fee(6.5%)', 'Seller Order Processing Fee(1250)', 'Premium(0.45%)']
            df_i_sub = df_i[income_fields].drop_duplicates('order number')
            df_m = pd.merge(df_s, df_i_sub, on='order number', how='left')

            # 匹配成本表 (关联 SKU 并匹配成本单价)
            # 注意：请确保成本表包含 'Nomor Referensi SKU' 和 '成本单价' 两列
            df_m = pd.merge(df_m, df_c[['Nomor Referensi SKU', '成本单价']], on='Nomor Referensi SKU', how='left')

            # 财务公式计算函数
            def calculate_row(row):
                if row['Status Pesanan'] == 'Batal': # 匹配印尼语取消状态
                    return 0, 0, 0, 0
                
                # 成功订单销售金额
                s_amt = row['Harga Setelah Diskon'] * row['Jumlah']
                # 优惠券分摊
                coupon = row['Voucher Ditanggung Penjual'] / row['计数'] if row['计数'] > 0 else 0
                # 平台费用汇总
                fees = sum([row.get(c, 0) for c in income_fields if c != 'order number'])
                # 最终收入
                inc = s_amt - coupon + fees
                # 成本计算
                c_total = row['成本单价'] * row['Jumlah'] if pd.notnull(row['成本单价']) else 0
                return s_amt, coupon, inc, c_total

            # 执行计算
            df_m[['成功订单销售金额', '优惠券', 'income', '成本']] = df_m.apply(lambda x: pd.Series(calculate_row(x)), axis=1)
            df_m['ad'] = "" # 预留广告费空列

            st.success("✅ 对账计算完成！")
            st.dataframe(df_m.head(10)) # 预览前10行

            # 3. 导出 Excel
            output = io.BytesIO()
            df_m.to_excel(output, index=False)
            st.download_button(
                label="📥 下载对账结果 Excel",
                data=output.getvalue(),
                file_name="Shopee_Final_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"处理失败，可能原因：Excel 表头不一致。具体错误: {e}")

# 提交并保存文件
