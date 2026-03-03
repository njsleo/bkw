import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import json
import io
import os
from duckduckgo_search import DDGS # 引入搜索引擎

# ==========================================
# 1. 页面与全局设置
# ==========================================
st.set_page_config(page_title="物理备课工作站", page_icon="⚛️", layout="wide", initial_sidebar_state="expanded")

try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except KeyError:
    st.error("⚠️ 未找到 API Key！请检查 Secrets 配置。")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

# 初始化 Session State (记忆库)
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "老师您好！我是您的物理教研助手。您可以直接和我探讨教学思路，或者点击下方按钮让我去全网搜索最新的引课素材，最后我可以在右侧为您生成 PPT。"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "source_text" not in st.session_state:
    st.session_state.source_text = ""

# ==========================================
# 2. 辅助功能函数
# ==========================================
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

# 联网搜索功能 (Manus 的核心绝招之一)
def search_web(query):
    try:
        results = DDGS().text(query, max_results=3)
        return "\n".join([f"- 【{r['title']}】: {r['body']}" for r in results])
    except Exception as e:
        return f"搜索失败: {e}"

# ==========================================
# 3. 界面布局：左、中、右三栏
# ==========================================

# 【左栏】：来源 (Sources) - 侧边栏
with st.sidebar:
    st.header("📚 来源 (Sources)")
    st.markdown("上传教材、大纲或试卷，我会牢记这些内容。")
    uploaded_file = st.file_uploader("添加来源文件...", type=['txt', 'docx'])
    
    if uploaded_file:
        st.session_state.source_text = read_file(uploaded_file)
        st.success(f"📄 已记忆: {uploaded_file.name}")
        with st.expander("查看来源内容"):
            st.write(st.session_state.source_text[:300] + "...")
            
    st.markdown("---")
    st.header("🎨 主题与模板")
    topic = st.text_input("当前备课课题：", placeholder="例如：光电效应")
    template_files = get_templates()
    selected_template = st.selectbox("选择 PPT 模板：", template_files) if template_files else None
    selected_template_path = os.path.join("templates", selected_template) if selected_template else None

# 主界面分为中右两栏 (比例 6:4)
col_chat, col_studio = st.columns([6, 4], gap="large")

# ==========================================
# 【中栏】：对话与 Agent 交互 (Chat)
# ==========================================
with col_chat:
    st.header("💬 智能教研对话")
    
    # 聊天记录展示区 (固定高度，类似真实聊天软件)
    chat_container = st.container(height=500)
    with chat_container:
        for msg in st.session_state.messages:
            st.chat_message(msg["role"]).write(msg["content"])

    # 快捷指令区 (模拟 Agent 自动执行任务)
    st.markdown("⚡ **快捷教研指令**")
    btn_col1, btn_col2 = st.columns(2)
    
    with btn_col1:
        if st.button("🌐 联网搜索前沿应用作引课", use_container_width=True):
            if not topic:
                st.warning("👈 请先在左侧输入【当前备课课题】")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请帮我联网搜索关于“{topic}”在真实世界、前沿科技中的最新应用，我想用来做课堂引课。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 正在全网检索最新科技新闻..."):
                            search_results = search_web(f"{topic} 最新科技应用 物理新闻")
                            
                            prompt = f"你是一个物理老师。用户想找关于“{topic}”的引课素材。我为你联网搜索到了以下最新资料：\n{search_results}\n\n请你根据这些资料，用幽默吸引人的口吻，写一段可以直接在课堂上说出来的“引课开场白”。"
                            
                            response = client.chat.completions.create(
                                model="deepseek-chat",
                                messages=[{"role": "user", "content": prompt}],
                                temperature=0.7
                            )
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col2:
        if st.button("🪄 一键生成幻灯片至 Studio", type="primary", use_container_width=True):
            if not topic:
                st.warning("👈 请先在左侧输入【当前备课课题】")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请根据我们刚才的聊天记录以及左侧的来源文档，为“{topic}”生成一份精美的 PPT，并发送到右侧的 Studio 中。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🎨 正在为您统筹排版，将数据发送至 Studio..."):
                            source_context = f"参考文档内容：\n{st.session_state.source_text}\n\n" if st.session_state.source_text else ""
                            chat_context = "之前的对话参考：\n" + "\n".join([m["content"] for m in st.session_state.messages[-3:]])
                            
                            prompt = f"""
                            {source_context}
                            {chat_context}
                            
                            请你结合以上资料，为课题“{topic}”生成 PPT 大纲。
                            【输出格式】必须是纯 JSON 数组格式，包含 image_suggestion（配图建议）：
                            [
                                {{"title": "幻灯片标题", "content": ["要点1", "要点2"], "image_suggestion": "配图要求"}}
                            ]
                            """
                            
                            try:
                                response = client.chat.completions.create(
                                    model="deepseek-chat",
                                    messages=[{"role": "user", "content": prompt}],
                                    temperature=0.3
                                )
                                result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                                
                                # 存入 session_state，右侧 Studio 会自动读取
                                st.session_state.ppt_data = json.loads(result_text)
                                
                                reply = "✅ 我已经为您提炼了核心内容，并将生成的幻灯片发送到了右侧的 **Studio** 工作区中，您可以直接预览并下载！"
                                st.write(reply)
                                st.session_state.messages.append({"role": "assistant", "content": reply})
                            except Exception as e:
                                st.error("生成失败，请重试。")

    # 底部自由聊天框
    if prompt := st.chat_input("输入其他教研问题 (如：这节课的重难点是什么？)"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.chat_message("user").write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    context_msg = [{"role": "system", "content": f"你是物理老师。当前资料：{st.session_state.source_text}"}] + st.session_state.messages
                    response = client.chat.completions.create(model="deepseek-chat", messages=context_msg, temperature=0.6)
                    reply = response.choices[0].message.content
                    st.write(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})

# ==========================================
# 【右栏】：Studio 工作区 (生成成果展示)
# ==========================================
with col_studio:
    st.header("✨ Studio 工作区")
    
    if st.session_state.ppt_data is None:
        # 空状态展示
        st.info("👈 左侧输入课题后，在中栏点击**「一键生成幻灯片」**，生成的 PPT 预览将出现在这里。")
        # 放一张占位图增加科技感
        st.image("https://images.unsplash.com/photo-1516321497487-e288fb19713f?q=80&w=2070&auto=format&fit=crop", caption="等待生成中...")
    else:
        # 渲染 PPT 文件
        if selected_template_path:
            prs = Presentation(selected_template_path)
        else:
            prs = Presentation() 

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
        st.download_button(
            label="📥 下载专属 PPT 文件",
            data=ppt_stream,
            file_name=f"{topic}_AI备课.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True
        )
        
        st.markdown("### 👀 逐页预览")
        with st.container(height=500): # 给预览区加个滚动条
            for i, slide_data in enumerate(st.session_state.ppt_data):
                with st.container(border=True):
                    st.markdown(f"**第 {i+1} 页：{slide_data.get('title', '无标题')}**")
                    for point in slide_data.get("content", []):
                        st.markdown(f"- {point}")
                    if slide_data.get("image_suggestion") and slide_data.get("image_suggestion") != "无":
                        st.info(f"🖼️ 配图：{slide_data.get('image_suggestion')}")