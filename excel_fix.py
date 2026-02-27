import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Excel 数据智能处理工具", layout="wide")
st.title("📊 Excel 数据智能处理工具 (Krazy)")

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

# ================= 界面划分为三个标签页 =================
tab1, tab2, tab3 = st.tabs(["🔗 功能一：跨表匹配合并", "🪄 功能二：单表内自动填补", "⚖️ 功能三：新旧版本智能比对"])

# ================= 【功能一：跨表匹配合并】 =================
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
            if st.checkbox("✅ 表1：自动向下填充空白数据", value=True, key='ffill1_tab1'):
                df1 = df1.ffill()
        with c_clean2:
            if st.checkbox("✅ 表2：自动向下填充空白数据", value=True, key='ffill2_tab1'):
                df2 = df2.ffill()

        st.divider()
        st.subheader("👀 3. 数据预览")
        pv_t1_1, pv_t1_2 = st.columns(2)
        with pv_t1_1:
            st.write("👉 **表1 预览 (前5行)**")
            st.dataframe(df1.head(5), use_container_width=True)
        with pv_t1_2:
            st.write("👉 **表2 预览 (前5行)**")
            st.dataframe(df2.head(5), use_container_width=True)

        st.divider()
        st.subheader("🔗 4. 匹配与提取设置")
        condition_count = st.number_input("匹配条件数量", min_value=1, max_value=10, value=4, key='cc1')
        left_on_cols, right_on_cols = [], []
        
        for i in range(int(condition_count)):
            c_a, c_b = st.columns(2)
            with c_a: left_on_cols.append(st.selectbox(f"条件 {i+1} : 表1 的列", df1.columns.tolist(), key=f"l_{i}"))
            with c_b: right_on_cols.append(st.selectbox(f"条件 {i+1} : 对应 表2 的列", df2.columns.tolist(), key=f"r_{i}"))

        available_targets = [c for c in df2.columns if c not in right_on_cols]
        target_cols = st.multiselect("选择你要从 表2 提取过来放到 表1 的列", available_targets, key='tc1')

        if left_on_cols and right_on_cols and target_cols:
            if st.button("🚀 开始跨表匹配合并", type="primary"):
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
                    st.download_button("📥 点击下载合并后的新文件", data=output, file_name="跨表匹配完成.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"匹配出错: {e}")

# ================= 【功能二：单表内自动填补】 =================
with tab2:
    st.subheader("📂 1. 上传包含残缺数据的表格")
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
            st.subheader("👀 2. 数据预览")
            st.write("👉 **原始数据预览 (前5行)**")
            st.dataframe(df_single.head(5), use_container_width=True)

            st.divider()
            st.subheader("🎯 3. 设置填补规则")
            c1, c2 = st.columns(2)
            with c1: match_cols_single = st.multiselect("1️⃣ 选择参照列 (用于查字典)", df_single.columns.tolist(), key='match_single')
            with c2: target_col_single = st.selectbox("2️⃣ 选择残缺列 (需要填补的目标)", [""] + df_single.columns.tolist(), key='target_single')
            
            if st.button("🚀 开始智能填补", type="primary", key="btn_single"):
                if not match_cols_single or target_col_single == "":
                    st.warning("⚠️ 请先选择参照列和目标残缺列！")
                else:
                    try:
                        df_res = df_single.copy()
                        df_res[target_col_single] = df_res[target_col_single].replace(r'^\s*$', np.nan, regex=True)
                        valid_data = df_res.dropna(subset=[target_col_single])
                        mapping_df = valid_data[match_cols_single + [target_col_single]].drop_duplicates(subset=match_cols_single)
                        mapping_df['__key__'] = mapping_df[match_cols_single].apply(tuple, axis=1)
                        mapping_dict = mapping_df.set_index('__key__')[target_col_single].to_dict()
                        
                        def fill_missing(row):
                            val = row[target_col_single]
                            if pd.isna(val): return mapping_dict.get(tuple(row[match_cols_single]), val)
                            return val
                        
                        df_res[target_col_single] = df_res.apply(fill_missing, axis=1)
                        st.success("✅ 填补完成！")
                        st.dataframe(df_res.head(10), use_container_width=True)
                        
                        output_single = io.BytesIO()
                        with pd.ExcelWriter(output_single, engine='openpyxl') as writer:
                            df_res.to_excel(writer, index=False, sheet_name='智能填补结果')
                        output_single.seek(0)
                        st.download_button("📥 点击下载填补完毕的新文件", data=output_single, file_name="表内自动填补完成.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"出错: {e}")

# ================= 【功能三：新旧版本智能比对】 =================
with tab3:
    st.subheader("📂 1. 上传新旧版本表格")
    col1, col2 = st.columns(2)
    with col1:
        file_old = st.file_uploader("📤 上传【旧版本】表格 (旧数据)", type=['xlsx', 'xls', 'csv'], key='f_old')
        df_old = None
        if file_old:
            h_old = st.number_input("⚙️ 旧版本 表头在第几行？", min_value=1, value=3, key='h_old') - 1
            if file_old.name.endswith(('xlsx', 'xls')):
                sheet_old = st.selectbox("选择 旧版本 Sheet", pd.ExcelFile(file_old).sheet_names, key='s_old')
                df_old = load_data(file_old, sheet_old, h_old)
            else:
                df_old = load_data(file_old, header_row=h_old)

    with col2:
        file_new = st.file_uploader("📤 上传【新版本】表格 (新数据)", type=['xlsx', 'xls', 'csv'], key='f_new')
        df_new = None
        if file_new:
            h_new = st.number_input("⚙️ 新版本 表头在第几行？", min_value=1, value=3, key='h_new') - 1
            if file_new.name.endswith(('xlsx', 'xls')):
                sheet_new = st.selectbox("选择 新版本 Sheet", pd.ExcelFile(file_new).sheet_names, key='s_new')
                df_new = load_data(file_new, sheet_new, h_new)
            else:
                df_new = load_data(file_new, header_row=h_new)

    if df_old is not None and df_new is not None:
        
        # --- 这里是为您全新加上的预览功能 ---
        st.divider()
        st.subheader("👀 2. 数据预览")
        pv_t3_1, pv_t3_2 = st.columns(2)
        with pv_t3_1:
            st.write("👉 **旧版本 预览 (前5行)**")
            st.dataframe(df_old.head(5), use_container_width=True)
        with pv_t3_2:
            st.write("👉 **新版本 预览 (前5行)**")
            st.dataframe(df_new.head(5), use_container_width=True)
        # ------------------------------------

        st.divider()
        st.subheader("🔗 3. 设置比对规则 (手动匹配列名)")
        st.info("💡 1. 唯一标识：相当于数据的“身份证”（如 订单号+款号），用来认出不管怎么乱序的同一行数据。\n\n💡 2. 待比对列：你需要检查是否被偷偷修改了内容的列（如 数量、价格、颜色）。")

        # --- 唯一标识列设置 ---
        key_count = st.number_input("你需要指定几个【唯一标识】列？", min_value=1, max_value=5, value=2, key='kc_diff')
        old_keys, new_keys = [], []
        for i in range(int(key_count)):
            c1, c2 = st.columns(2)
            with c1: old_keys.append(st.selectbox(f"🔑 旧表 - 唯一标识 {i+1}", df_old.columns.tolist(), key=f"ok_{i}"))
            with c2: new_keys.append(st.selectbox(f"🔑 新表 - 对应标识 {i+1}", df_new.columns.tolist(), key=f"nk_{i}"))

        st.divider()
        
        # --- 待比对内容列设置 ---
        cmp_count = st.number_input("你需要比对几个【内容】列？", min_value=1, max_value=15, value=1, key='cc_diff')
        old_cmps, new_cmps = [], []
        for i in range(int(cmp_count)):
            c1, c2 = st.columns(2)
            with c1: old_cmps.append(st.selectbox(f"🔍 旧表 - 待比对列 {i+1}", df_old.columns.tolist(), key=f"oc_{i}"))
            with c2: new_cmps.append(st.selectbox(f"🔍 新表 - 对应比对列 {i+1}", df_new.columns.tolist(), key=f"nc_{i}"))

        if st.button("🚀 扫描差异并生成比对报告", type="primary"):
            try:
                # 提取需要的列并统一列名为“新表”的列名，方便合并对比
                df_old_sub = df_old[old_keys + old_cmps].copy()
                df_old_sub.columns = new_keys + new_cmps
                # 去重（防止多行相同主键导致爆炸）
                df_old_sub = df_old_sub.drop_duplicates(subset=new_keys)
                
                df_new_sub = df_new[new_keys + new_cmps].copy()
                df_new_sub = df_new_sub.drop_duplicates(subset=new_keys)

                # 将所有数据转为字符串，消除空值带来的对比误差
                df_old_sub = df_old_sub.astype(str).replace('nan', '')
                df_new_sub = df_new_sub.astype(str).replace('nan', '')

                # 以外连接的方式合并
                merged = pd.merge(df_old_sub, df_new_sub, on=new_keys, how='outer', suffixes=('_旧版', '_新版'), indicator=True)

                # 1. 找出被删除的行 (只在旧表出现)
                deleted_df = merged[merged['_merge'] == 'left_only'][new_keys + [c + '_旧版' for c in new_cmps]]
                deleted_df.columns = new_keys + new_cmps
                
                # 2. 找出新增的行 (只在新表出现)
                added_df = merged[merged['_merge'] == 'right_only'][new_keys + [c + '_新版' for c in new_cmps]]
                added_df.columns = new_keys + new_cmps

                # 3. 找出共有的行，并检查是否被修改
                both_df = merged[merged['_merge'] == 'both'].copy()
                
                details = []
                for idx, row in both_df.iterrows():
                    diffs = []
                    for c in new_cmps:
                        old_v = str(row[c + '_旧版']).strip()
                        new_v = str(row[c + '_新版']).strip()
                        if old_v != new_v:
                            diffs.append(f"【{c}】由 '{old_v}' 变更为 '{new_v}'")
                    if diffs:
                        details.append("； ".join(diffs))
                    else:
                        details.append("")
                
                both_df['修改详情'] = details
                # 过滤出真正被修改的行
                modified_df = both_df[both_df['修改详情'] != ""]
                
                # 整理一下修改表的列显示顺序
                show_cols = new_keys.copy()
                for c in new_cmps:
                    show_cols.extend([c + '_旧版', c + '_新版'])
                show_cols.append('修改详情')
                modified_df = modified_df[show_cols]

                st.success(f"✅ 扫描完成！发现：新增 {len(added_df)} 行，删除 {len(deleted_df)} 行，修改 {len(modified_df)} 行。")

                # 在网页上展示三个状态的表格
                st.write("🟢 **新增的数据 (仅在新版有)**")
                st.dataframe(added_df.head(5), use_container_width=True)
                
                st.write("🔴 **删除的数据 (仅在旧版有)**")
                st.dataframe(deleted_df.head(5), use_container_width=True)
                
                st.write("🟡 **被修改的数据 (红色为旧，绿色为新)**")
                st.dataframe(modified_df.head(5), use_container_width=True)

                # 生成拥有三个 Sheet 的专业 Excel 报告
                output_diff = io.BytesIO()
                with pd.ExcelWriter(output_diff, engine='openpyxl') as writer:
                    modified_df.to_excel(writer, index=False, sheet_name='修改的数据')
                    added_df.to_excel(writer, index=False, sheet_name='新增的数据')
                    deleted_df.to_excel(writer, index=False, sheet_name='删除的数据')
                output_diff.seek(0)
                
                st.download_button("📥 下载完整比对报告 (含 修改、新增、删除 3个Sheet)", data=output_diff, file_name="新旧版本比对报告.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

            except Exception as e:
                st.error(f"比对出错，请检查是否有两边数据类型不一致等问题。错误详情: {e}")

