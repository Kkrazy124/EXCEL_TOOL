import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Excel 数据智能处理工具", layout="wide")
st.title("📊 Excel 数据智能处理工具 (Krazy)")

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

# ================= 界面划分为两个标签页 =================
tab1, tab2 = st.tabs(["🔗 功能一：跨表匹配合并 (双表/多Sheet)", "🪄 功能二：单表内自动填补 (补全缺失数据)"])

# ================= 【功能一：跨表匹配合并】 (保留原有功能) =================
with tab1:
    mode = st.radio("请选择上传模式：", ("上传两个独立的文件", "上传一个文件 (包含多个Sheet)"), key="mode_tab1")
    df1, df2 = None, None

    st.subheader("📂 1. 上传与读取设置")
    if mode == "上传两个独立的文件":
        col1, col2 = st.columns(2)
        with col1:
            file1 = st.file_uploader("📤 上传 表1 (基础表)", type=['xlsx', 'xls', 'csv'], key='f1')
            if file1:
                h1 = st.number_input("⚙️ 表1 表头在第几行？", min_value=1, value=3, key='h1') - 1
                if file1.name.endswith(('xlsx', 'xls')):
                    sheet1 = st.selectbox("选择 Sheet (表1)", pd.ExcelFile(file1).sheet_names, key='s1')
                    df1 = load_data(file1, sheet1, h1)
                else:
                    df1 = load_data(file1, header_row=h1)
        with col2:
            file2 = st.file_uploader("📤 上传 表2 (数据源)", type=['xlsx', 'xls', 'csv'], key='f2')
            if file2:
                h2 = st.number_input("⚙️ 表2 表头在第几行？", min_value=1, value=3, key='h2') - 1
                if file2.name.endswith(('xlsx', 'xls')):
                    sheet2 = st.selectbox("选择 Sheet (表2)", pd.ExcelFile(file2).sheet_names, key='s2')
                    df2 = load_data(file2, sheet2, h2)
                else:
                    df2 = load_data(file2, header_row=h2)
    else:
        file = st.file_uploader("📤 上传包含多个 Sheet 的 Excel 文件", type=['xlsx', 'xls'], key='f_multi')
        if file:
            excel_file = pd.ExcelFile(file)
            sheets = excel_file.sheet_names
            if len(sheets) < 2:
                st.error("该 Excel 文件中只有 1 个 Sheet，无法进行表间匹配！")
            else:
                col1, col2 = st.columns(2)
                with col1:
                    sheet1 = st.selectbox("选择 表1 (基础表)", sheets, index=0)
                    h1 = st.number_input("⚙️ 表1 表头在第几行？", min_value=1, value=3, key='h1_s') - 1
                    df1 = load_data(file, sheet1, h1)
                with col2:
                    sheet2 = st.selectbox("选择 表2 (数据源)", sheets, index=1)
                    h2 = st.number_input("⚙️ 表2 表头在第几行？", min_value=1, value=3, key='h2_s') - 1
                    df2 = load_data(file, sheet2, h2)

    if df1 is not None and df2 is not None:
        st.divider()
        st.subheader("🧹 2. 数据清洗 (解决合并单元格空白)")
        c_clean1, c_clean2 = st.columns(2)
        with c_clean1:
            if st.checkbox("✅ 表1：自动向下填充空白数据", value=True, key='ffill1'):
                df1 = df1.ffill()
        with c_clean2:
            if st.checkbox("✅ 表2：自动向下填充空白数据", value=True, key='ffill2'):
                df2 = df2.ffill()

        st.divider()
        st.subheader("🔗 3. 匹配与提取设置")
        condition_count = st.number_input("匹配条件数量", min_value=1, max_value=10, value=4)
        left_on_cols, right_on_cols = [], []
        
        for i in range(int(condition_count)):
            col_a, col_b = st.columns(2)
            with col_a:
                sel_1 = st.selectbox(f"条件 {i+1} : 表1 的列", df1.columns.tolist(), key=f"l_{i}")
                left_on_cols.append(sel_1)
            with col_b:
                sel_2 = st.selectbox(f"条件 {i+1} : 对应 表2 的列", df2.columns.tolist(), key=f"r_{i}")
                right_on_cols.append(sel_2)

        available_targets = [c for c in df2.columns if c not in right_on_cols]
        target_cols = st.multiselect("选择你要从 表2 提取过来放到 表1 的列", available_targets)

        if left_on_cols and right_on_cols and target_cols:
            if st.button("🚀 开始双表匹配合并", type="primary"):
                try:
                    df2_subset = df2[right_on_cols + target_cols].drop_duplicates(subset=right_on_cols)
                    result_df = pd.merge(df1, df2_subset, left_on=left_on_cols, right_on=right_on_cols, how='left')
                    cols_to_drop = [col for col in right_on_cols if col not in left_on_cols and col in result_df.columns]
                    if cols_to_drop: result_df = result_df.drop(columns=cols_to_drop)
                    
                    st.success("✅ 匹配成功！")
                    st.dataframe(result_df.head(10), use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='匹配结果')
                    output.seek(0)
                    st.download_button("📥 点击下载合并后的新 Excel 文件", data=output, file_name="跨表匹配完成.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"匹配出错: {e}")

# ================= 【功能二：单表内自动填补】 (全新功能) =================
with tab2:
    st.subheader("📂 1. 上传包含残缺数据的表格")
    st.info("💡 **原理说明**：工具会自动从表中找出完整的数据作为“词典”。对于下方带有空白的残缺数据，只要你选定的匹配列对得上，就会自动把上面的内容复制下来填入空白处。")
    
    file_single = st.file_uploader("📤 上传表格", type=['xlsx', 'xls', 'csv'], key='f_single')
    if file_single:
        h_single = st.number_input("⚙️ 表头在第几行？", min_value=1, value=3, key='h_single') - 1
        
        df_single = None
        if file_single.name.endswith(('xlsx', 'xls')):
            sheet_single = st.selectbox("选择要处理的 Sheet", pd.ExcelFile(file_single).sheet_names, key='s_single')
            df_single = load_data(file_single, sheet_single, h_single)
        else:
            df_single = load_data(file_single, header_row=h_single)

        if df_single is not None:
            st.divider()
            st.write("👉 **原始数据预览 (请确认哪些列是用来匹配的，哪一列是残缺需要填的)**")
            st.dataframe(df_single.head(5), use_container_width=True)
            
            st.divider()
            st.subheader("🎯 2. 设置填补规则")
            c1, c2 = st.columns(2)
            with c1:
                match_cols_single = st.multiselect(
                    "1️⃣ 选择用于匹配的【参照列】 (支持多选，如：订单号、款号)", 
                    df_single.columns.tolist(), 
                    key='match_single'
                )
            with c2:
                target_col_single = st.selectbox(
                    "2️⃣ 选择需要自动填补的【目标残缺列】 (如：中文颜色名)", 
                    [""] + df_single.columns.tolist(), 
                    key='target_single'
                )
            
            if st.button("🚀 开始表内智能填补", type="primary", key="btn_single"):
                if not match_cols_single or target_col_single == "":
                    st.warning("⚠️ 请先选择参照列和目标残缺列！")
                else:
                    try:
                        df_res = df_single.copy()
                        
                        # 1. 把看起来是空的单元格统一变成标准缺失值 (NaN)
                        df_res[target_col_single] = df_res[target_col_single].replace(r'^\s*$', np.nan, regex=True)
                        
                        # 2. 提取出目标列有值的行，作为我们的“词典参照库”
                        valid_data = df_res.dropna(subset=[target_col_single])
                        
                        # 3. 建立映射字典 (去重，保留第一条有效规则)
                        mapping_df = valid_data[match_cols_single + [target_col_single]].drop_duplicates(subset=match_cols_single)
                        # 将多列条件合并为元组作为键
                        mapping_df['__key__'] = mapping_df[match_cols_single].apply(tuple, axis=1)
                        mapping_dict = mapping_df.set_index('__key__')[target_col_single].to_dict()
                        
                        # 4. 定义逐行填补的函数
                        def fill_missing(row):
                            val = row[target_col_single]
                            # 如果当前行是空的，就去字典里查
                            if pd.isna(val):
                                key = tuple(row[match_cols_single])
                                return mapping_dict.get(key, val) # 查不到就保持原样
                            return val # 如果本来就有值，就不动它
                        
                        # 5. 执行填补
                        df_res[target_col_single] = df_res.apply(fill_missing, axis=1)
                        
                        st.success("✅ 表内空白填补完成！预览一下看看效果：")
                        st.dataframe(df_res.head(15), use_container_width=True)
                        
                        # 6. 生成下载文件
                        output_single = io.BytesIO()
                        with pd.ExcelWriter(output_single, engine='openpyxl') as writer:
                            df_res.to_excel(writer, index=False, sheet_name='智能填补结果')
                        output_single.seek(0)
                        
                        st.download_button(
                            label="📥 点击下载填补完毕的新 Excel 文件",
                            data=output_single,
                            file_name="表内自动填补完成.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"填补过程中出现错误: {e}")
