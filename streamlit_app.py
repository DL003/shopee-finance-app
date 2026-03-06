import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Shopee 财务对账助手", layout="wide")
st.title("🚀 Shopee 财务账单自动化核算 (逻辑重构版)")

def get_col(df, keywords):
    """精准定位列名：去除空格、转小写后包含关键字即可"""
    for col in df.columns:
        c_clean = str(col).strip().lower()
        for k in keywords:
            if k.lower() in c_clean:
                return col
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
    if st.button("🚀 开始自动化对账"):
        try:
            # 1. 加载数据
            df_a = pd.read_excel(f_a)
            df_b = pd.read_excel(f_b)
            df_c = pd.read_excel(f_c)
            df_d = pd.read_excel(f_d)

            # 2. 识别订单号
            oid_b = get_col(df_b, ['order number', '订单号'])
            oid_c = get_col(df_c, ['order number', '订单号'])
            
            if not oid_b or not oid_c:
                st.error("❌ 无法识别订单号列，请检查表B和表C。")
                st.stop()

            # 3. 提取表C费用逻辑 (强制转数值，避免空值报错)
            fee_map = {
                'ams': ['AMS Commission Fee'],
                'comm': ['Commission fee'],
                'service': ['Service Fee'],
                'proc': ['Processing Fee', 'Seller Order Processing'],
                'premium': ['Premium']
            }
            
            # 创建干净的收入明细表
            df_c_clean = df_c[[oid_c]].copy()
            found_fee_internal_names = []
            
            for key, keywords in fee_map.items():
                actual_col = get_col(df_c, keywords)
                if actual_col:
                    df_c_clean[key] = pd.to_numeric(df_c[actual_col], errors='coerce').fillna(0)
                    found_fee_internal_names.append(key)
                else:
                    df_c_clean[key] = 0.0
            
            # 收入表去重
            df_c_final = df_c_clean.drop_duplicates(oid_c)

            # 4. 准备成本数据
            sku_b = get_col(df_b, ['Nomor Referensi SKU', 'SKU'])
            sku_d = get_col(df_d, ['Nomor Referensi SKU', 'SKU'])
            price_d = get_col(df_d, ['成本', '单价'])

            # 5. 合并数据流
            df_b['计数'] = df_b.groupby(oid_b)[oid_b].transform('count')
            df_main = pd.merge(df_b, df_c_final, left_on=oid_b, right_on=oid_c, how='left')
            
            if sku_b and sku_d and price_d:
                df_main = pd.merge(df_main, df_d[[sku_d, price_d]], left_on=sku_b, right_on=sku_d, how='left')

            # 6. 执行核心财务计算
            status_col = get_col(df_b, ['Status Pesanan', '订单状态'])
            sale_price_col = get_col(df_b, ['Harga Setelah Diskon', '折后价'])
            qty_col = get_col(df_b, ['Jumlah', '数量'])
            voucher_col = get_col(df_b, ['Voucher Ditanggung Penjual', '优惠券'])

            df_main = df_main.fillna(0)

            def calc_logic(row):
                # 状态检查
                st_str = str(row.get(status_col, '')).lower()
                if any(x in st_str for x in ['batal', 'cancelled', '取消']):
                    return 0.0, 0.0, 0.0, 0.0
                
                # 原始值获取
                s_price = row.get(sale_price_col, 0)
                qty = row.get(qty_col, 0)
                v_total = row.get(voucher_col, 0)
                cnt = row.get('计数', 1)
                
                # 计算中间项
                sales_amt = s_price * qty
                v_share = v_total / cnt if cnt > 0 else 0
                
                # 费用汇总 (直接从我们定义的 key 中取)
                fee_sum = row['ams'] + row['comm'] + row['service'] + row['proc'] + row['premium']
                
                income = sales_amt - v_share + fee_sum
                cost = row.get(price_d, 0) * qty
                
                return sales_amt, v_share, income, cost

            df_main[['_S', '_V', '_I', '_C']] = df_main.apply(lambda r: pd.Series(calc_logic(r)), axis=1)

            # 7. 映射回表 A 模板
            df_result = pd.DataFrame(columns=df_a.columns)
            
            for col in df_result.columns:
                c_name = str(col).strip()
                # A. 费用明细映射
                matched_fee = False
                for key, keywords in fee_map.items():
                    if keywords[0].lower() in c_name.lower():
                        df_result[col] = df_main[key]
                        matched_fee = True
                        break
                if matched_fee: continue

                # B. 计算结果映射
                if '成功订单销售金额' in c_name: df_result[col] = df_main['_S']
                elif 'income' in c_name.lower(): df_result[col] = df_main['_I']
                elif '最终成本' in c_name or ('成本' in c_name and '单价' not in c_name): df_result[col] = df_main['_C']
                elif '优惠券' in c_name: df_result[col] = df_main['_V']
                else:
                    # C. 原始字段映射
                    orig = get_col(df_main, [c_name])
                    if orig: df_result[col] = df_main[orig]

            st.success("✅ 核算成功！已生成最终结果。")
            st.dataframe(df_result.head(10))

            # 导出
            out = io.BytesIO()
            df_result.to_excel(out, index=False)
            st.download_button("📥 下载 Shopee 核算结果", out.getvalue(), "Shopee_Final_Report.xlsx")

        except Exception as e:
            st.error(f"❌ 运行报错: {str(e)}")
