import streamlit as st
import pandas as pd
import io
import warnings
from datetime import datetime

# 忽略 Excel 样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 缓存函数：只有当文件内容改变时才会重新读取 ---
@st.cache_data(show_spinner="正在极速解析 Excel 数据...")
def load_excel(file):
    if file is None:
        return None
    return pd.read_excel(file)

# 初始化 Session State (用于存储历史记录)
if 'history' not in st.session_state:
    st.session_state['history'] = []

st.set_page_config(page_title="Excel 动态合并工具", layout="wide")

# --- 顶部导航 ---
col_title, col_reset = st.columns([8, 1])
with col_title:
    st.title("🚀 Excel 动态合并助手")
with col_reset:
    if st.button("🔄 重置页面"):
        st.cache_data.clear()  # 清除缓存
        st.rerun()

# --- 1. 文件上传区 ---
st.header("第 1 步：上传 Excel 文件")
col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("上传【文件 1】(主表 / 商家表)", type=['xlsx', 'xls'])
with col2:
    file2 = st.file_uploader("上传【文件 2】(数据源 / Tracking 表)", type=['xlsx', 'xls'])

if file1 and file2:
    # 使用缓存读取数据
    df1 = load_excel(file1)
    df2 = load_excel(file2)

    st.divider()

    # --- 2. 使用 Form 减少刷新频率 ---
    with st.form("merge_config_form"):
        st.header("第 2 步：配置合并逻辑与字段")
        
        c1, c2 = st.columns(2)
        with c1:
            key1 = st.selectbox("文件 1 的关联列", options=df1.columns)
        with c2:
            key2 = st.selectbox("文件 2 的关联列", options=df2.columns)

        source_columns = [col for col in df2.columns if col != key2]
        selected_cols = st.multiselect("请选择要从 文件 2 提取的列:", options=source_columns)
        
        # 表单提交按钮
        submit_button = st.form_submit_button(label='🔥 执行合并')

    # --- 3. 处理合并逻辑 ---
    if submit_button:
        if not selected_cols:
            st.warning("⚠️ 请至少勾选一列需要提取的数据。")
        else:
            try:
                # 处理数据
                df1_proc = df1.copy()
                df2_proc = df2.copy()
                
                # 统一转为字符串并去空格
                df1_proc[key1] = df1_proc[key1].astype(str).str.strip()
                df2_proc[key2] = df2_proc[key2].astype(str).str.strip()

                # 提取选中的列
                df2_subset = df2_proc[[key2] + selected_cols]
                df2_subset = df2_subset.rename(columns={key2: key1})

                # 合并数据
                result_df = pd.merge(df1_proc, df2_subset, on=key1, how='left')

                # 更新历史记录
                now = datetime.now().strftime("%H:%M:%S")
                st.session_state['history'].insert(0, {
                    "时间": now,
                    "操作": f"合并了 {len(selected_cols)} 列数据",
                    "总行数": len(result_df)
                })

                st.success("✅ 合并成功！")
                st.subheader("合并结果预览 (前 10 行)")
                st.dataframe(result_df.head(10))

                # 下载区域
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="💾 点击下载合并后的 Excel",
                    data=output.getvalue(),
                    file_name=f"Merged_Result_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"合并过程中出现错误: {e}")

# --- 4. 历史记录 (仅限本次会话) ---
if st.session_state['history']:
    st.divider()
    with st.expander("查看本次操作历史记录"):
        st.table(pd.DataFrame(st.session_state['history']))
else:
    st.info("💡 提示：上传并配置好字段后，点击“执行合并”即可开始。")