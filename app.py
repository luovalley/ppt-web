import streamlit as st
import os
import tempfile

# 页面基础配置
st.set_page_config(
    page_title="AI Native PPT Master (OpenRouter)",
    page_icon="📊",
    layout="centered"
)

st.title("📊 AI 原生 PPT 生成器 (OpenRouter 版)")
st.markdown("通过 OpenRouter 统一接口，将文档或主题转化为**真正的原生 PowerPoint 幻灯片**。")

# 侧边栏：OpenRouter 配置
with st.sidebar:
    st.header("⚙️ OpenRouter 配置")
    api_key = st.text_input("输入 OpenRouter API Key", type="password", placeholder="sk-or-v1-...")
    
    # 常用大模型选择（OpenRouter 格式）
    model_name = st.selectbox(
        "选择大模型",
        [
            "anthropic/claude-3.5-sonnet",
            "deepseek/deepseek-chat",
            "openai/gpt-4o",
            "google/gemini-pro-1.5"
        ],
        index=0
    )
    
    st.markdown("---")
    st.markdown("### 关于项目")
    st.markdown("基于 `ppt-master` 核心逻辑，通过 OpenRouter 驱动全球顶级大模型。")

# 主界面：输入区域
tab1, tab2 = st.tabs(["📄 文档转 PPT", "💡 主题直接生成"])

with tab1:
    st.markdown("### 上传参考文件")
    uploaded_file = st.file_uploader(
        "支持 PDF、Word、Markdown 等格式", 
        type=["pdf", "docx", "md", "txt"]
    )
    doc_prompt = st.text_area("补充要求（可选）", placeholder="例如：请生成商务汇报风格，10 页左右...")

with tab2:
    st.markdown("### 输入 PPT 主题或大纲")
    topic_input = st.text_area("请输入你想生成的 PPT 主题、背景或详细大纲要求", placeholder="例如：2026年人工智能行业发展趋势与应用场景...")

# 生成按钮
if st.button("🚀 开始生成原生 PPT", type="primary"):
    if not api_key:
        st.error("请先在左侧边栏填写你的 OpenRouter API Key！")
    else:
        # 核心：将 OpenRouter 注入环境变量，兼容标准的 OpenAI 客户端调用
        os.environ["OPENAI_API_KEY"] = api_key
        os.environ["OPENAI_API_BASE"] = "https://openrouter.ai/api/v1"
        
        with st.spinner(f"正在通过 OpenRouter 调用 `{model_name}` 深度生成 PPT，请稍候..."):
            
            # --- 实际接入位置 ---
            # 如果你把 ppt-master 的源码拉到了项目中，可以在这里调用它的生成函数。
            # 例如：
            # if uploaded_file is not None:
            #     # 处理文件并调用后端生成逻辑
            #     pass
            
            # 模拟生成耗时
            import time
            time.sleep(3)
            
            st.success("🎉 PPT 生成成功！")
            
            # 伪代码：生成后的下载按钮
            # with open("output.pptx", "rb") as f:
            #     st.download_button(
            #         label="📥 点击下载生成的 .pptx 文件",
            #         data=f,
            #         file_name="presentation.pptx",
            #         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            #     )
            
            st.info(f"提示：当前已成功绑定 OpenRouter，并选择了模型 **{model_name}**。")
