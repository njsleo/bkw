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
# 1. 页面与全局设置
# ==========================================
st.set_page_config(page_title="物理教研工作站", page_icon="⚛️", layout="wide", initial_sidebar_state="expanded")

try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except KeyError:
    st.error("⚠️ 未找到 API Key！请检查 Secrets 配置。")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "老师好！我已经切换至**纯本地安全模式**。您可以在左侧一次性选择多份电脑里的资料，我会仔细阅读它们，并随时为您提供硬核的物理学史引课！"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "current_context" not in st.session_state:
    st.session_state.current_context = ""

# ==========================================
# 2. 辅助功能函数
# ==========================================
def read_file(file_obj):
    try:
        if file_obj.name.endswith('.txt'):
            return file_obj.getvalue().decode("utf-8")
        elif file_obj.name.endswith('.docx'):
            return "\n".join([para.text for para in docx.Document(file_obj).paragraphs])
    except Exception as e:
        return f"读取 {file_obj.name} 失败: {e}"
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

# 新增功能：将 AI 生成的 JSON 数据转化为 Word 文档
def generate_word_document(ppt_data, topic_name):
    doc = docx.Document()
    # 大标题
    doc.add_heading(f"【教学讲义】{topic_name}", 0)
    
    # 遍历每一页的内容写入 Word
    for i, slide in enumerate(ppt_data):
        doc.add_heading(f"环节 {i+1}：{slide.get('title', '无标题')}", level=2)
        for point in slide.get("content", []):
            doc.add_paragraph(point, style='List Bullet')
        
        # 加上配图建议
        img_sug = slide.get("image_suggestion", "")
        if img_sug and img_sug != "无":
            p = doc.add_paragraph(f"🖼️ [配图建议：{img_sug}]")
            p.runs[0].font.italic = True
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# ==========================================
# 3. 界面布局：左、中、右
# ==========================================

with st.sidebar:
    st.header("📂 本地资料读取")
    st.markdown("资料不上传云端，仅在本次对话中读取，绝对安全。")
    
    uploaded_files = st.file_uploader(
        "请选择电脑里的教案/试题 (可按住 Command 键多选)：", 
        type=['txt', 'docx'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("📥 一键读取选中的所有文件", use_container_width=True):
            with st.spinner("正在快速阅读您选中的本地文件..."):
                all_content = []
                for file in uploaded_files:
                    file_text = read_file(file)
                    all_content.append(f"--- 来自本地文件《{file.name}》 ---\n{file_text}")
                
                st.session_state.current_context = "\n\n".join(all_content)
                st.success(f"✅ 成功读取了 {len(uploaded_files)} 份文档！")
                with st.expander("查看已读取的内容概览"):
                    st.write(st.session_state.current_context[:500] + "......")

    st.markdown("---")
    st.header("🎨 主题与模板")
    topic = st.text_input("当前备课课题：", placeholder="例如：法拉第电磁感应定律")
    template_files = get_templates()
    selected_template = st.selectbox("选择 PPT 模板：", template_files) if template_files else None
    selected_template_path = os.path.join("templates", selected_template) if selected_template else None

col_chat, col_studio = st.columns([6, 4], gap="large")

with col_chat:
    st.header("💬 智能教研对话")
    
    chat_container = st.container(height=500)
    with chat_container:
        for msg in st.session_state.messages:
            st.chat_message(msg["role"]).write(msg["content"])

    st.markdown("⚡ **快捷教研指令**")
    btn_col1, btn_col2, btn_col3 = st.columns(3)
    
    with btn_col1:
        if st.button("🕰️ 找真实物理学史作引课", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请以真实客观的科学史视角，帮我挖掘“{topic}”的发现历程和真实历史细节，用来做硬核的课堂引课。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 正在检索深度的物理史料与真实文献..."):
                            search_results = search_web(f"{topic} 物理学史 真实历史事件 科学家传记 维基百科")
                            prompt = f"""
                            你是一位严谨的【科学史学者】和资深的高中物理特级教师。
                            我需要你为课题“{topic}”准备一段具有【真实历史厚重感】的课堂引入。
                            以下是我为你联网检索到的真实历史资料：\n{search_results}\n
                            要求：
                            1. 绝对拒绝低幼、虚构的童话式口吻（严禁出现“从前有个科学家”、“他灵机一动”这种轻浮表达）。
                            2. 必须基于真实的物理学史实。要具体到【真实的年份】、【当时的物理学界认知背景/瓶颈】以及【科学家面临的真实困境或实验误差】。
                            3. 语言风格要类似《典籍里的中国》或 BBC 科学纪录片的旁白，要有深度、有悬念，用科学的严谨来激发高中生的智力好奇心。
                            4. 如果检索资料中有科学家的真实名言或著作原话，请务必准确引用。
                            """
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.current_context += "\n\n" + reply 
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col2:
        if st.button("📄 从本地文档中提取", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            elif not st.session_state.current_context: st.warning("请先在左侧选择并读取您的本地电脑文档！")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请结合我刚才上传的本地文档，提取关于“{topic}”的备课核心要点。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🧠 正在梳理多份文档的内容..."):
                            prompt = f"这是我提供的多份本地资料内容汇总：\n{st.session_state.current_context}\n\n请你从中找出与“{topic}”最相关的内容，并整理成严谨的备课要点。"
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col3:
        if st.button("🪄 一键排版输出", type="primary", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请根据我们刚聊的内容，为“{topic}”排版成大纲，发送到Studio。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🎨 正在为您统筹排版..."):
                            chat_history = "\n".join([m["content"] for m in st.session_state.messages[-3:]])
                            prompt = f"""
                            参考资料：{chat_history}
                            请结合参考资料，为课题“{topic}”生成大纲。确保把【物理学史/科学家故事】放在前几页作为引入。
                            【输出格式】必须是纯 JSON 数组，包含 image_suggestion：
                            [ {{"title": "标题", "content": ["要点1", "要点2"], "image_suggestion": "配图建议"}} ]
                            """
                            try:
                                response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.2)
                                result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                                st.session_state.ppt_data = json.loads(result_text)
                                reply = "✅ 讲义大纲已发送至右侧 Studio，您现在可以下载 PPT 和 Word 讲义了！"
                                st.write(reply)
                                st.session_state.messages.append({"role": "assistant", "content": reply})
                            except Exception as e:
                                st.error("生成失败，请重试。")

    if prompt := st.chat_input("自由对话..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.chat_message("user").write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    context_msg = [{"role": "system", "content": "你是资深物理老师。"}] + st.session_state.messages
                    response = client.chat.completions.create(model="deepseek-chat", messages=context_msg, temperature=0.6)
                    reply = response.choices[0].message.content
                    st.write(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})

with col_studio:
    st.header("✨ Studio 工作区")
    if st.session_state.ppt_data is None:
        st.info("👈 在中栏点击生成按钮后，成果预览将出现在这里。")
    else:
        # 1. 准备 PPT 字节流
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
        
        # 2. 准备 Word 字节流
        word_stream = generate_word_document(st.session_state.ppt_data, topic if topic else "未命名课题")
        
        # 3. 提供双格式下载按钮
        st.success("🎉 生成成功！请选择您需要的格式进行下载：")
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                "📊 下载 PPT 演示版", 
                data=ppt_stream, 
                file_name=f"{topic}_AI备课.pptx", 
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
                type="primary", 
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                "📝 下载 Word 教案版", 
                data=word_stream, 
                file_name=f"{topic}_教案讲义.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                type="secondary", 
                use_container_width=True
            )
        
        st.markdown("### 👀 讲义内容预览")
        with st.container(height=500):
            for i, slide_data in enumerate(st.session_state.ppt_data):
                with st.container(border=True):
                    st.markdown(f"**环节 {i+1}：{slide_data.get('title', '无标题')}**")
                    for point in slide_data.get("content", []): st.markdown(f"- {point}")
                    if slide_data.get("image_suggestion") and slide_data.get("image_suggestion") != "无": st.info(f"🖼️ 配图：{slide_data.get('image_suggestion')}")