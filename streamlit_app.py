import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务核算助手", layout="wide")
st.title("🛍️ Shopee 账单自动化核算 (基于真实表格适配版)")

# 智能匹配函数：忽略大小写和空格
def get_col_exact(df, target_name):
    target_clean = str(target_name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target_clean:
            return col
    # 如果没找到完全一样的，尝试模糊匹配
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
    if st.button("🚀 开始自动化对账"):
        try:
            # 读取数据 (增加容错，处理CSV或Excel)
            def load_df(file):
                try:
                    return pd.read_excel(file)
                except:
                    return pd.read_csv(file)

            df_temp = load_df(f_a)
            df_order = load_df(f_b)
            df_income = load_df(f_c)
            df_cost = load_df(f_d)

            # 1. 寻找核心关联键
            oid_order = get_col_exact(df_order, 'order number')
            oid_income = get_col_exact(df_income, 'Order number')
            sku_order = get_col_exact(df_order, 'Nomor Referensi SKU')
            sku_cost = get_col_exact(df_cost, 'Nomor Referensi SKU')
            cost_val = get_col_exact(df_cost, '成本单价')

            if not oid_order or not oid_income:
                st.error("❌ 无法匹配订单号列，请检查订单表和收入表的表头。")
                st.stop()

            # 2. 预处理收入表 (提取5项费用)
            # 这里的关键字完全对应你上传的 "订单收入.xlsx"
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
                    df_income_clean[key] = pd.to_numeric(df_income[actual_col], errors='coerce').fillna(0)
                else:
                    df_income_clean[key] = 0.0
            
            df_income_final = df_income_clean.drop_duplicates(oid_income)

            # 3. 关联数据
            # 计算计数
            df_order['计数'] = df_order.groupby(oid_order)[oid_order].transform('count')
            
            # 合并收入
            df_main = pd.merge(df_order, df_income_final, left_on=oid_order, right_on=oid_income, how='left')
            
            # 合并成本
            if sku_order and sku_cost and cost_val:
                df_main = pd.merge(df_main, df_cost[[sku_cost, cost_val]], left_on=sku_order, right_on=sku_cost, how='left')
            else:
                df_main['成本单价'] = 0.0

            # 4. 执行财务公式
            # 对应订单表.xlsx的原始列名
            status_col = get_col_exact(df_order, 'Status Pesanan')
            price_col = get_col_exact(df_order, 'Harga Setelah Diskon')
            qty_col = get_col_exact(df_order, 'Jumlah')
            voucher_col = get_col_exact(df_order, 'Voucher Ditanggung Penjual')

            df_main = df_main.fillna(0)

            def calc_row(row):
                # 状态判断
                st_str = str(row.get(status_col, '')).lower()
                if 'batal' in st_str or 'cancel' in st_str:
                    return 0.0, 0.0, 0.0, 0.0
                
                # 销售额 = 折后价 * 数量
                s_amt = row.get(price_col, 0) * row.get(qty_col, 0)
                # 优惠券平摊
                v_share = row.get(voucher_col, 0) / row['计数'] if row['计数'] > 0 else 0
                # 费用汇总
                fees = row['ams'] + row['comm'] + row['service'] + row['proc'] + row['premium']
                # Income
                inc = s_amt - v_share + fees
                # 单行成本
                c_total = row.get(cost_val, 0) * row.get(qty_col, 0)
                return s_amt, v_share, inc, c_total

            df_main[['_S', '_V', '_I', '_C']] = df_main.apply(lambda r: pd.Series(calc_row(r)), axis=1)

            # 5. 映射回模板 A
            # 这里的逻辑：只要模板里的列名在 df_main 里能找到，就填进去
            df_final = pd.DataFrame(columns=df_temp.columns)
            
            # 定义一个映射字典，处理特殊计算项
            special_map = {
                '成功订单销售金额': '_S',
                '优惠券': '_V',
                'income': '_I',
                '成本': cost_val,
                '总成本': '_C'
            }

            for col in df_final.columns:
                c_name = str(col).strip()
                # A. 先查费用明细映射
                matched_fee = False
                for k_internal, k_orig in income_fee_map.items():
                    if k_orig.lower() in c_name.lower():
                        df_final[col] = df_main[k_internal]
                        matched_fee = True
                        break
                if matched_fee: continue

                # B. 查特殊计算结果
                if c_name in special_map:
                    df_final[col] = df_main[special_map[c_name]]
                else:
                    # C. 查原始字段
                    orig = get_col_exact(df_main, c_name)
                    if orig:
                        df_final[col] = df_main[orig]

            st.success("✅ 对账完成！请预览并下载结果。")
            st.dataframe(df_final.head(10))

            # 导出
            out = io.BytesIO()
            df_final.to_excel(out, index=False)
            st.download_button("📥 下载 Shopee 核算结果", out.getvalue(), "Shopee_Final_A.xlsx")

        except Exception as e:
            st.error(f"❌ 逻辑执行报错: {str(e)}")
