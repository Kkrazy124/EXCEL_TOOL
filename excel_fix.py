import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel 数据智能匹配合并工具", layout="wide")
st.title("📊 Excel/CSV 数据智能匹配工具 (Krazy)")

# 选择上传模式
mode = st.radio("请选择上传模式：", ("上传两个独立的文件", "上传一个文件 (包含多个Sheet)"))

# 缓存读取数据的函数
@st.cache_data
def load_data(file, sheet_name=None, header_row=0):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, header=header_row)
        else:
            return pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    except Exception as e:
        st.error(f"读取文件出错，请确认表头行数是否正确。错误信息: {e}")
        return None

df1, df2 = None, None

st.divider()
st.subheader("📂 1. 上传与读取设置")

if mode == "上传两个独立的文件":
    col1, col2 = st.columns(2)
    with col1:
        file1 = st.file_uploader("📤 上传 表1 (需要填入新数据的基础表)", type=['xlsx', 'xls', 'csv'], key='f1')
        if file1:
            h1 = st.number_input("⚙️ 表1 的表头在第几行？", min_value=1, value=3, key='h1') - 1
            if file1.name.endswith(('xlsx', 'xls')):
                sheet1 = st.selectbox("选择 Sheet (表1)", pd.ExcelFile(file1).sheet_names, key='s1')
                df1 = load_data(file1, sheet1, h1)
            else:
                df1 = load_data(file1, header_row=h1)
            
    with col2:
        file2 = st.file_uploader("📤 上传 表2 (包含新数据的数据源表)", type=['xlsx', 'xls', 'csv'], key='f2')
        if file2:
            h2 = st.number_input("⚙️ 表2 的表头在第几行？", min_value=1, value=3, key='h2') - 1
            if file2.name.endswith(('xlsx', 'xls')):
                sheet2 = st.selectbox("选择 Sheet (表2)", pd.ExcelFile(file2).sheet_names, key='s2')
                df2 = load_data(file2, sheet2, h2)
            else:
                df2 = load_data(file2, header_row=h2)

else:
    file = st.file_uploader("📤 上传包含多个 Sheet 的 Excel 文件", type=['xlsx', 'xls'])
    if file:
        excel_file = pd.ExcelFile(file)
        sheets = excel_file.sheet_names
        if len(sheets) < 2:
            st.error("该 Excel 文件中只有 1 个 Sheet，无法进行表间匹配！")
        else:
            col1, col2 = st.columns(2)
            with col1:
                sheet1 = st.selectbox("选择 表1 (基础表)", sheets, index=0)
                h1 = st.number_input("⚙️ 表1 的表头在第几行？", min_value=1, value=3, key='h1_s') - 1
                df1 = load_data(file, sheet1, h1)
            with col2:
                sheet2 = st.selectbox("选择 表2 (数据源表)", sheets, index=1)
                h2 = st.number_input("⚙️ 表2 的表头在第几行？", min_value=1, value=3, key='h2_s') - 1
                df2 = load_data(file, sheet2, h2)

# 如果两张表都已成功读取
if df1 is not None and df2 is not None:
    st.divider()
    st.subheader("🧹 2. 数据清洗 (解决合并单元格空白)")
    
    # 增加自动向下填充的功能
    c_clean1, c_clean2 = st.columns(2)
    with c_clean1:
        st.info("💡 如果 表1 中有合并单元格导致读取出空白(NaN)，请勾选下方：")
        if st.checkbox("✅ 表1：自动向下填充空白数据", value=True, key='ffill1'):
            df1 = df1.ffill()
    with c_clean2:
        st.info("💡 如果 表2 中有合并单元格导致读取出空白(NaN)，请勾选下方：")
        if st.checkbox("✅ 表2：自动向下填充空白数据", value=True, key='ffill2'):
            df2 = df2.ffill()

    st.divider()
    st.subheader("👀 3. 数据预览 (请核对空白是否已被填充)")
    c1, c2 = st.columns(2)
    with c1:
        st.write("👉 **表1 预览 (前5行)**")
        st.dataframe(df1.head(5), use_container_width=True)
    with c2:
        st.write("👉 **表2 预览 (前5行)**")
        st.dataframe(df2.head(5), use_container_width=True)

    st.divider()
    st.subheader("🔗 4. 对应匹配条件设置")
    
    condition_count = st.number_input("你需要几个条件来确认唯一匹配？(例如：需要订单号、款号、颜色、数量 4个)", min_value=1, max_value=10, value=4)
    
    left_on_cols = []
    right_on_cols = []
    
    for i in range(int(condition_count)):
        col_a, col_b = st.columns(2)
        with col_a:
            default_index_1 = i if i < len(df1.columns) else 0
            sel_1 = st.selectbox(f"条件 {i+1} : 表1 的列", df1.columns.tolist(), index=default_index_1, key=f"l_{i}")
            left_on_cols.append(sel_1)
        with col_b:
            default_index_2 = i if i < len(df2.columns) else 0
            sel_2 = st.selectbox(f"条件 {i+1} : 对应 表2 的列", df2.columns.tolist(), index=default_index_2, key=f"r_{i}")
            right_on_cols.append(sel_2)

    st.divider()
    st.subheader("🎯 5. 提取目标数据")
    available_targets = [c for c in df2.columns if c not in right_on_cols]
    target_cols = st.multiselect("请选择你要从 表2 提取过来放到 表1 的列 (如：中文颜色名)", available_targets)

    if left_on_cols and right_on_cols and target_cols:
        if st.button("🚀 开始匹配并生成新表", type="primary"):
            try:
                # 提取并去重
                df2_subset = df2[right_on_cols + target_cols].drop_duplicates(subset=right_on_cols)
                
                # 执行合并
                result_df = pd.merge(df1, df2_subset, left_on=left_on_cols, right_on=right_on_cols, how='left')
                
                # 清理冗余列
                cols_to_drop = [col for col in right_on_cols if col not in left_on_cols and col in result_df.columns]
                if cols_to_drop:
                    result_df = result_df.drop(columns=cols_to_drop)
                
                st.success("✅ 匹配成功！预览前 10 行：")
                st.dataframe(result_df.head(10), use_container_width=True)
                
                # 导出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='匹配结果')
                output.seek(0)
                
                st.download_button(
                    label="📥 点击下载合并后的新 Excel 文件",
                    data=output,
                    file_name="匹配完成_新文件.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"匹配出错，请检查数据。错误详情: {e}")