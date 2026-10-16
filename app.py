import streamlit as st
import os
import tempfile

# 页面基础配置
st.set_page_config(
    page_title="AI Native PPT Master (OpenRouter)",
    page_icon="📊",
    layout="centered"
)

# OpenRouter 默认模型与免费备用模型配置
DEFAULT_OPENROUTER_MODEL = os.getenv(
    "OPENROUTER_MODEL", "nvidia/nemotron-3-super-120b-a12b:free"
)
FALLBACK_FREE_MODELS = [
    "nvidia/nemotron-3-super-120b-a12b:free",
    "qwen/qwen3.8-27b:free",
    "google/gemma-4-31b-it:free",
    "apodex/apodex-1.1-mini:free",
    "openrouter/free",
]

st.title("📊 AI 原生 PPT 生成器 (OpenRouter 版)")
st.markdown("通过 OpenRouter 统一接口，将文档或主题转化为**真正的原生 PowerPoint 幻灯片**。")

# 侧边栏：OpenRouter 配置
with st.sidebar:
    st.header("⚙️ OpenRouter 配置")
    api_key = st.text_input("输入 OpenRouter API Key", type="password", placeholder="sk-or-v1-...")
    
    # 组合模型列表（将免费/备用模型与常规主力模型合并）
    all_models = FALLBACK_FREE_MODELS + [
        "anthropic/claude-3.5-sonnet",
        "deepseek/deepseek-chat",
        "openai/gpt-4o",
        "google/gemini-pro-1.5"
    ]
    
    # 确定默认选项的索引
    default_index = all_models.index(DEFAULT_OPENROUTER_MODEL) if DEFAULT_OPENROUTER_MODEL in all_models else 0
    
    model_name = st.selectbox(
        "选择大模型",
        all_models,
        index=default_index
    )
    
    st.markdown("---")
    st.markdown("### 关于项目")
    st.markdown("基于 `ppt-master` 核心逻辑，通过 OpenRouter 驱动大模型。")

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
        # 核心：将 OpenRouter 注入环境变量与当前选择的模型
        os.environ["OPENAI_API_KEY"] = api_key
        os.environ["OPENAI_API_BASE"] = "https://openrouter.ai/api/v1"
        os.environ["OPENROUTER_MODEL"] = model_name
        
        with st.spinner(f"正在通过 OpenRouter 调用模型 `{model_name}` 深度生成 PPT，请稍候..."):
            
            # --- 实际接入位置 ---
            # 后续可在此处引入 ppt-master 核心脚本进行文件解析与 PPT 渲染
            
            import time
            time.sleep(3)
            
            st.success("🎉 PPT 生成成功！")
            
            # 伪代码：下载按钮
            # with open("output.pptx", "rb") as f:
            #     st.download_button(
            #         label="📥 点击下载生成的 .pptx 文件",
            #         data=f,
            #         file_name="presentation.pptx",
            #         mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            #     )
            
            st.info(f"提示：当前已绑定 OpenRouter，并选择了模型 **{model_name}**。")
