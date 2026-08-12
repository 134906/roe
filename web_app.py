import streamlit as st
import sys
import os
import tempfile
import pandas as pd
from core import run_analysis

# 设置页面配置
st.set_page_config(
    page_title="ROE SCAN 自动化分析",
    page_icon="🌿",
    layout="wide"
)

# ---------- 自定义 CSS 样式（绿色主题） ----------
st.markdown("""
<style>
    .stApp {
        background-color: #E8F5E9;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #1B5E20;
    }
    .css-1d391kg, .css-1d391kg .st-emotion-cache-1wmy9hl {
        background-color: #2E7D32;
    }
    .css-1d391kg .st-emotion-cache-1wmy9hl, .css-1d391kg label, .css-1d391kg .st-emotion-cache-1wmy9hl p {
        color: #FFFFFF;
    }
    .css-1d391kg .st-emotion-cache-1wmy9hl h1, .css-1d391kg .st-emotion-cache-1wmy9hl h2, .css-1d391kg .st-emotion-cache-1wmy9hl h3 {
        color: #E8F5E9;
    }
    .stButton > button {
        background-color: #388E3C;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 2rem;
        font-weight: bold;
    }
    .stButton > button:hover {
        background-color: #2E7D32;
        color: white;
        border: 1px solid #A5D6A7;
    }
    .stFileUploader > div {
        border: 2px dashed #4CAF50;
        border-radius: 10px;
        background-color: #C8E6C9;
    }
    .stTextArea > div {
        background-color: #FFFFFF;
        border-radius: 8px;
        border: 1px solid #A5D6A7;
    }
    .stAlert {
        border-radius: 8px;
    }
    .stDownloadButton > button {
        background-color: #1B5E20;
        color: white;
    }
    .stDownloadButton > button:hover {
        background-color: #0D3B0E;
    }
    p, li, label, .stMarkdown {
        color: #1B5E20;
    }
    .css-1d391kg p, .css-1d391kg label, .css-1d391kg .stMarkdown {
        color: #FFFFFF;
    }
    .css-1d391kg a {
        color: #A5D6A7;
    }
</style>
""", unsafe_allow_html=True)

# ---------- 页面标题 ----------
st.title("ROE SCAN 自动化分析")
st.markdown("上传 SPSS 数据文件和配置文件，点击运行即可获得完整分析结果。")

# ---------- 侧边栏：文件上传与参数 ----------
with st.sidebar:
    st.header("📁 文件上传")
    spss_file = st.file_uploader("上传 SPSS 数据文件 (.sav)", type=["sav"])
    config_file = st.file_uploader("上传配置文件 (.xlsx)", type=["xlsx"])
    workfile_file = st.file_uploader("（可选）上传 Workfile 模板 (.xlsx)", type=["xlsx"])
    output_filename = st.text_input("输出文件名（不含路径）", value="ROE_Results.xlsx")

# ---------- 主区域：运行按钮 ----------
run_btn = st.button("运行", disabled=(spss_file is None or config_file is None))

# ---------- 初始化会话状态 ----------
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []
if 'run_completed' not in st.session_state:
    st.session_state.run_completed = False

# 日志显示区域
log_area = st.empty()

def update_log_display():
    """刷新日志显示区域"""
    if st.session_state.log_messages:
        # 显示全部日志（可滚动）
        log_area.text_area("📝 运行日志", "\n".join(st.session_state.log_messages), height=400)
    else:
        log_area.text_area("📝 运行日志", "等待运行...", height=400)

# 初始显示
update_log_display()

# ---------- 运行逻辑 ----------
if run_btn:
    # 清空旧日志并重置完成状态
    st.session_state.log_messages = []
    st.session_state.run_completed = False
    update_log_display()  # 立即清空显示

    if spss_file is None or config_file is None:
        st.error("请先上传 SPSS 数据和配置文件！")
    else:
        # 保存上传文件到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_spss:
            tmp_spss.write(spss_file.getvalue())
            spss_path = tmp_spss.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_config:
            tmp_config.write(config_file.getvalue())
            config_path = tmp_config.name

        # 输出结果文件路径
        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name

        # 处理 Workfile
        workfile_path = None
        if workfile_file is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_wf:
                tmp_wf.write(workfile_file.getvalue())
                workfile_path = tmp_wf.name

        # 自定义日志回调（追加到 session_state）
        def log_callback(msg):
            st.session_state.log_messages.append(msg)
            update_log_display()  # 每次更新显示

        # 执行分析
        try:
            run_analysis(spss_path, config_path, output_path,
                         workfile_path=workfile_path,
                         log_callback=log_callback)
            st.session_state.run_completed = True
            st.success("✅ 分析完成！点击下方按钮下载结果。")

            # 读取主结果文件提供下载
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 下载主结果 Excel",
                    data=f,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 如果提供了 workfile，则提供填充后的 workfile 下载
            if workfile_path is not None:
                base, ext = os.path.splitext(output_filename)
                workfile_out_name = f"{base}_workfile_filled{ext}"
                workfile_out_path = os.path.join(os.path.dirname(output_path), workfile_out_name)
                if os.path.exists(workfile_out_path):
                    with open(workfile_out_path, "rb") as f:
                        st.download_button(
                            label="📥 下载 Workfile 填充结果",
                            data=f,
                            file_name=workfile_out_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

        except Exception as e:
            st.error(f"❌ 分析失败: {e}")
            log_callback(str(e))
        finally:
            # 清理临时文件
            try:
                os.unlink(spss_path)
                os.unlink(config_path)
                os.unlink(output_path)
                if workfile_path is not None:
                    os.unlink(workfile_path)
                    # 可能的 workfile_out_path 也会被清理，但为了安全，保留
            except:
                pass

# 如果已经运行完成，但用户尚未重新运行，下载按钮会保留（因为它们在 run_btn 块内，
# 但 run_btn 只会在点击后触发，之后这些组件仍然显示在界面上，直到下次点击运行被清空）