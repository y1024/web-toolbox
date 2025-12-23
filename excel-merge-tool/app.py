# 本地启动 streamlit run app.py
import streamlit as st
import pandas as pd
import io
import warnings
from datetime import datetime

# 忽略 Excel 样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 初始化 Session State (用于存储历史记录)
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.set_page_config(page_title="Excel 动态合并工具", layout="wide")

# --- 顶部标题和重置功能 ---
col_title, col_reset = st.columns([8, 1])
with col_title:
    st.title("🚀 Excel 动态合并助手")
with col_reset:
    # 利用 Streamlit 的 rerun 机制实现重置
    if st.button("🔄 重置"):
        st.rerun()

# --- 1. 文件上传区 ---
st.header("第 1 步：上传 Excel 文件")
col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("上传【文件 1】(主表)", type=['xlsx', 'xls'], key="u1")
with col2:
    file2 = st.file_uploader("上传【文件 2】(数据源)", type=['xlsx', 'xls'], key="u2")

if file1 and file2:
    try:
        # 使用缓存读取数据，避免重复加载
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        
        st.divider()

        # --- 2. 配置字段关联 ---
        st.header("第 2 步：配置关联逻辑")
        c1, c2 = st.columns(2)
        with c1:
            key1 = st.selectbox("文件 1 的关联列", options=df1.columns)
        with c2:
            key2 = st.selectbox("文件 2 的关联列", options=df2.columns)

        # --- 3. 选择要搬运的列 ---
        st.header("第 3 步：选择要合并的列")
        source_columns = [col for col in df2.columns if col != key2]
        selected_cols = st.multiselect("请选择要从 文件 2 提取的列:", options=source_columns)

        if selected_cols:
            if st.button("🔥 执行合并并生成预览"):
                # 数据处理
                df1_proc = df1.copy()
                df2_proc = df2.copy()
                df1_proc[key1] = df1_proc[key1].astype(str).str.strip()
                df2_proc[key2] = df2_proc[key2].astype(str).str.strip()

                df2_subset = df2_proc[[key2] + selected_cols]
                df2_subset = df2_subset.rename(columns={key2: key1})

                # 合并
                result_df = pd.merge(df1_proc, df2_subset, on=key1, how='left')

                # 记录到历史 (Session State)
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                history_item = {
                    "时间": now,
                    "主表名": file1.name,
                    "来源表": file2.name,
                    "合并列数": len(selected_cols),
                    "总行数": len(result_df)
                }
                st.session_state['history'].insert(0, history_item) # 新记录排在前面

                st.success("✅ 合并完成！")
                st.dataframe(result_df.head(10))

                # 下载
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="💾 点击下载合并后的 Excel",
                    data=output.getvalue(),
                    file_name=f"已合并_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"处理出错: {e}")

# --- 4. 历史记录显示区 ---
st.divider()
with st.expander("查看本次操作历史记录", expanded=False):
    if st.session_state['history']:
        history_df = pd.DataFrame(st.session_state['history'])
        st.table(history_df)
    else:
        st.write("暂无记录")