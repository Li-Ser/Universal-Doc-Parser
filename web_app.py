import streamlit as st
import os
import tempfile
import shutil
import time
import zipfile
from datetime import datetime

# 导入核心解析引擎
# 确保 src/__init__.py 存在
try:
    from src.main import process_file_full_suite, scout_file
except ImportError:
    st.error("❌ 无法导入 src.main。请确保目录下有 src/main.py 且 src/__init__.py 存在。")
    st.stop()

# ===================================
# 页面配置 (Page Config)
# ===================================
st.set_page_config(
    page_title="Universal Doc Parser",
    page_icon="🌌",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义 CSS 美化
st.markdown("""
<style>
    .main-header {font-size: 2.5rem; color: #4F8BF9; font-weight: 700;}
    .sub-header {font-size: 1.2rem; color: #666;}
    .stButton>button {width: 100%; border-radius: 5px; height: 3em; font-weight: bold;}
    .success-box {padding: 1rem; background-color: #d4edda; border-radius: 5px; color: #155724;}
</style>
""", unsafe_allow_html=True)

# ===================================
# 侧边栏配置 (Sidebar)
# ===================================
with st.sidebar:
    st.title("⚙️ 导出配置")
    st.write("选择你需要的数据格式：")
    
    fmt_xlsx = st.checkbox("Excel (.xlsx)", value=True)
    fmt_json = st.checkbox("JSON (.json)", value=True)
    fmt_docx = st.checkbox("Word (.docx)", value=False)
    fmt_pdf  = st.checkbox("PDF (.pdf) [Beta]", value=False)
    
    formats = []
    if fmt_xlsx: formats.append('xlsx')
    if fmt_json: formats.append('json')
    if fmt_docx: formats.append('docx')
    if fmt_pdf:  formats.append('pdf')
    
    st.markdown("---")
    st.info("💡 **提示**：PDF 导出需要服务器安装中文字体，否则可能跳过。")
    st.markdown("---")
    st.markdown("### 关于项目")
    st.markdown("开源 RAG 文档清洗引擎。\n支持目录树、正则、视觉等多引擎智能调度。")

# ===================================
# 主界面 (Main UI)
# ===================================
st.markdown('<div class="main-header">🌌 Universal Doc Parser</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">RAG 知识库构建 / 智能文档分块 / 表格重构</div>', unsafe_allow_html=True)
st.markdown("---")

# 1. 文件上传区
uploaded_files = st.file_uploader(
    "📄 拖入 PDF 或 Word 文档 (支持批量)", 
    type=["pdf", "docx"], 
    accept_multiple_files=True
)

# 2. 处理逻辑
if uploaded_files:
    st.write(f"已加载 {len(uploaded_files)} 个文件。准备就绪。")
    
    start_btn = st.button(f"🚀 开始解析 ({len(uploaded_files)} 个文件)")
    
    if start_btn:
        if not formats:
            st.error("请至少在左侧选择一种导出格式！")
            st.stop()

        # 创建临时目录用于处理
        with tempfile.TemporaryDirectory() as temp_input_dir:
            with tempfile.TemporaryDirectory() as temp_output_dir:
                
                # 进度条
                progress_bar = st.progress(0)
                status_text = st.empty()
                logs_expander = st.expander("查看详细处理日志", expanded=True)
                
                processed_count = 0
                
                # 遍历处理文件
                for i, uploaded_file in enumerate(uploaded_files):
                    file_name = uploaded_file.name
                    status_text.text(f"正在处理: {file_name} ...")
                    
                    # 1. 保存上传的文件到临时目录 (解析器需要物理路径)
                    file_path = os.path.join(temp_input_dir, file_name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # 2. 调用核心引擎
                    with logs_expander:
                        st.write(f"**[{i+1}/{len(uploaded_files)}] 开始分析: {file_name}**")
                        try:
                            # 捕获 print 输出其实比较麻烦，这里我们简单显示“智能侦察中”
                            # 如果想做得更细，可以改造 main.py 返回 log
                            engine = scout_file(file_path)
                            st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;🕵️ 智能路由侦测: `{engine}`")
                            
                            process_file_full_suite(
                                file_path, 
                                temp_output_dir, 
                                formats
                            )
                            st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;✅ 解析完成")
                        except Exception as e:
                            st.error(f"❌ 处理失败: {e}")
                    
                    # 更新进度
                    processed_count += 1
                    progress_bar.progress(processed_count / len(uploaded_files))
                
                status_text.text("🎉 所有任务处理完成！正在打包...")
                
                # 3. 打包结果为 ZIP
                zip_name = f"parsed_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                zip_path = os.path.join(tempfile.gettempdir(), zip_name)
                
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(temp_output_dir):
                        for file in files:
                            zipf.write(
                                os.path.join(root, file), 
                                arcname=file
                            )
                
                # 4. 显示下载按钮
                with open(zip_path, "rb") as f:
                    st.markdown('<div class="success-box">✅ 解析成功！请点击下方按钮下载结果。</div>', unsafe_allow_html=True)
                    st.download_button(
                        label="📦 下载所有结果 (ZIP)",
                        data=f,
                        file_name=zip_name,
                        mime="application/zip"
                    )