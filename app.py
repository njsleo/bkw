import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import json
import io
import os
from duckduckgo_search import DDGS

# ==========================================
# 新增：RAG 知识库核心组件
# ==========================================
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_core.documents import Document

st.set_page_config(page_title="物理备课工作站", page_icon="⚛️", layout="wide", initial_sidebar_state="expanded")

try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except KeyError:
    st.error("⚠️ 未找到 API Key！请检查 Secrets 配置。")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

# ==========================================
# 初始化 AI 的“海马体” (向量嵌入模型)
# ==========================================
@st.cache_resource
def get_embeddings():
    # 使用专门针对中文优化的轻量级向量模型，体积小且免费
    return HuggingFaceEmbeddings(model_name="BAAI/bge-small-zh-v1.5")

embeddings = get_embeddings()
DB_PATH = "faiss_index"

def load_db():
    if os.path.exists(DB_PATH):
        return FAISS.load_local(DB_PATH, embeddings, allow_dangerous_deserialization=True)
    return None

# 初始化状态
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "老师好！我已经升级了**永久记忆引擎**。在左侧上传教案并存入知识库后，以后无论何时，您都可以随时让我去记忆库里翻找资料！"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "current_context" not in st.session_state:
    st.session_state.current_context = ""

# 辅助函数
def read_file(uploaded_file):
    if uploaded_file.name.endswith('.txt'):
        return uploaded_file.getvalue().decode("utf-8")
    elif uploaded_file.name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        return "\n".join([para.text for para in doc.paragraphs])
    return ""

def get_templates():
    if not os.path.exists("templates"): return []
    return [f for f in os.listdir("templates") if f.endswith(".pptx")]

def search_web(query):
    try:
        results = DDGS().text(query, max_results=3)
        return "\n".join([f"- 【{r['title']}】: {r['body']}" for r in results])
    except Exception as e:
        return f"搜索失败: {e}"

# ==========================================
# 【左栏】：知识库管理 (Knowledge Base)
# ==========================================
with st.sidebar:
    st.header("🧠 个人物理知识库")
    st.markdown("将历史教案、好题本存入此库，AI 将永久记住它们。")
    
    uploaded_file = st.file_uploader("添加新资料到知识库：", type=['txt', 'docx'])
    
    if uploaded_file and st.button("📥 提取并永久存入知识库", use_container_width=True):
        with st.spinner("正在将文档切片并转化为向量记忆..."):
            text = read_file(uploaded_file)
            # 1. 文本切片
            text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
            chunks = text_splitter.split_text(text)
            docs = [Document(page_content=chunk, metadata={"source": uploaded_file.name}) for chunk in chunks]
            
            # 2. 存入向量数据库 (FAISS)
            db = load_db()
            if db:
                db.add_documents(docs)
            else:
                db = FAISS.from_documents(docs, embeddings)
            
            # 3. 保存到本地文件夹
            db.save_local(DB_PATH)
            st.success(f"✅ 成功将《{uploaded_file.name}》切分为 {len(chunks)} 个记忆碎片并存入大脑！")

    st.markdown("---")
    st.header("🎨 主题与模板")
    topic = st.text_input("当前备课课题：", placeholder="例如：电磁感应")
    template_files = get_templates()
    selected_template = st.selectbox("选择 PPT 模板：", template_files) if template_files else None
    selected_template_path = os.path.join("templates", selected_template) if selected_template else None

col_chat, col_studio = st.columns([6, 4], gap="large")

# ==========================================
# 【中栏】：聊天区
# ==========================================
with col_chat:
    st.header("💬 智能教研对话")
    
    chat_container = st.container(height=500)
    with chat_container:
        for msg in st.session_state.messages:
            st.chat_message(msg["role"]).write(msg["content"])

    st.markdown("⚡ **快捷教研指令**")
    btn_col1, btn_col2, btn_col3 = st.columns(3)
    
    with btn_col1:
        if st.button("🌐 联网找引课", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"帮我联网找“{topic}”的最新科技应用作引课。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 检索中..."):
                            search_results = search_web(f"{topic} 科技应用")
                            prompt = f"根据以下最新资料，用幽默口吻写一段引课开场白：\n{search_results}"
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.7)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.current_context = reply # 暂存为生成PPT的上下文
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col2:
        if st.button("🧠 从知识库翻找", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                db = load_db()
                if not db:
                    st.error("知识库还是空的，请先在左侧上传资料存入大脑！")
                else:
                    st.session_state.messages.append({"role": "user", "content": f"请从我的个人知识库中，翻找关于“{topic}”的历史备课资料，并提取最核心的要点。"})
                    with chat_container:
                        st.chat_message("user").write(st.session_state.messages[-1]["content"])
                        with st.chat_message("assistant"):
                            with st.spinner("🧠 正在大脑深处检索记忆..."):
                                # 核心：向量相似度检索
                                results = db.similarity_search(topic, k=4)
                                kb_content = "\n\n".join([f"【来自片段 {i+1}】: {r.page_content}" for i, r in enumerate(results)])
                                
                                prompt = f"以下是从我的历史知识库中检索到的关于“{topic}”的记忆碎片：\n{kb_content}\n\n请帮我整理这些内容，形成一份清晰的备课要点。"
                                response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3)
                                reply = response.choices[0].message.content
                                st.write(reply)
                                st.session_state.current_context = reply 
                                st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col3:
        if st.button("🪄 一键做成 PPT", type="primary", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请根据我们刚聊的内容，为“{topic}”生成PPT并发送到Studio。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🎨 排版中..."):
                            prompt = f"""
                            参考背景信息：{st.session_state.current_context}
                            
                            请结合背景信息，为课题“{topic}”生成 PPT 大纲。
                            【输出格式】必须是纯 JSON 数组，包含 image_suggestion：
                            [ {{"title": "标题", "content": ["要点1", "要点2"], "image_suggestion": "配图建议"}} ]
                            """
                            try:
                                response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.2)
                                result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                                st.session_state.ppt_data = json.loads(result_text)
                                reply = "✅ 幻灯片已发送至右侧 Studio！"
                                st.write(reply)
                                st.session_state.messages.append({"role": "assistant", "content": reply})
                            except Exception as e:
                                st.error("生成失败。")

    if prompt := st.chat_input("自由对话..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.chat_message("user").write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    context_msg = [{"role": "system", "content": "你是物理老师。"}] + st.session_state.messages
                    response = client.chat.completions.create(model="deepseek-chat", messages=context_msg, temperature=0.6)
                    reply = response.choices[0].message.content
                    st.write(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})

# ==========================================
# 【右栏】：Studio
# ==========================================
with col_studio:
    st.header("✨ Studio 工作区")
    if st.session_state.ppt_data is None:
        st.info("👈 在中栏点击生成按钮后，PPT 预览将出现在这里。")
    else:
        if selected_template_path: prs = Presentation(selected_template_path)
        else: prs = Presentation() 

        for slide_data in st.session_state.ppt_data:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            if slide.shapes.title: slide.shapes.title.text = slide_data.get("title", "无标题")
            if len(slide.placeholders) > 1:
                tf = slide.placeholders[1].text_frame
                contents = slide_data.get("content", [""])
                if contents:
                    tf.text = contents[0]
                    for point in contents[1:]:
                        p = tf.add_paragraph()
                        p.text = point
                        p.level = 0
                
                img_sug = slide_data.get("image_suggestion", "")
                if img_sug and img_sug != "无":
                    p_img = tf.add_paragraph()
                    p_img.text = f"\n[🖼️ 配图建议：{img_sug}]"
                    p_img.level = 0
                    p_img.font.color.rgb = RGBColor(0, 112, 192)

        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        
        st.success("🎉 PPT 生成成功！")
        st.download_button("📥 下载专属 PPT 文件", data=ppt_stream, file_name=f"{topic}_AI备课.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
        
        st.markdown("### 👀 逐页预览")
        with st.container(height=500):
            for i, slide_data in enumerate(st.session_state.ppt_data):
                with st.container(border=True):
                    st.markdown(f"**第 {i+1} 页：{slide_data.get('title', '无标题')}**")
                    for point in slide_data.get("content", []): st.markdown(f"- {point}")
                    if slide_data.get("image_suggestion") and slide_data.get("image_suggestion") != "无": st.info(f"🖼️ 配图：{slide_data.get('image_suggestion')}")