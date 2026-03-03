import streamlit as st
from openai import OpenAI
from zhipuai import ZhipuAI
import docx
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import json
import io
import os
import requests
from duckduckgo_search import DDGS

# ==========================================
# 1. 页面与全局设置
# ==========================================
st.set_page_config(page_title="物理教研 Agent", page_icon="⚛️", layout="wide", initial_sidebar_state="expanded")

try:
    api_key_brain = st.secrets["DEEPSEEK_API_KEY"]
    api_key_paint = st.secrets["ZHIPU_API_KEY"]
except KeyError:
    st.error("⚠️ 未找到 API Key！请检查 .streamlit/secrets.toml 配置。")
    st.stop()

# 初始化双引擎
client_brain = OpenAI(api_key=api_key_brain, base_url="https://api.deepseek.com")
client_paint = ZhipuAI(api_key=api_key_paint)

if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "老师好！我是满血版的物理 Agent。我现在不仅能提炼教案、讲历史，还能**召唤智谱 CogView 为您画出专属物理配图**，并自动排版进 PPT 和 Word 中！"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "current_context" not in st.session_state:
    st.session_state.current_context = ""

# ==========================================
# 2. 核心技能引擎
# ==========================================
def read_file(file_obj):
    try:
        if file_obj.name.endswith('.txt'):
            return file_obj.getvalue().decode("utf-8")
        elif file_obj.name.endswith('.docx'):
            return "\n".join([para.text for para in docx.Document(file_obj).paragraphs])
    except Exception:
        return ""
    return ""

def get_templates():
    if not os.path.exists("templates"): return []
    return [f for f in os.listdir("templates") if f.endswith(".pptx")]

def search_web(query):
    try:
        results = DDGS().text(query, max_results=3)
        return "\n".join([f"- 【{r['title']}】: {r['body']}" for r in results])
    except Exception:
        return "搜索失败"

def generate_physics_image(image_prompt):
    """调用智谱大模型作图"""
    try:
        response = client_paint.images.generations(
            model="cogview-3-plus", 
            prompt=f"一张专业、严谨的高中物理教材插图，用于教学演示。画面内容：{image_prompt}。要求：纯白背景，线条清晰，不要出现任何错乱的英文字母或乱码文字，纯粹展示物理原理或实验现象。",
        )
        image_url = response.data[0].url
        img_data = requests.get(image_url).content
        return img_data
    except Exception as e:
        return None

def generate_word_document(ppt_data, topic_name):
    doc = docx.Document()
    
    # 设置全局仿宋
    styles_to_change = ['Normal', 'List Bullet', 'Title', 'Heading 1', 'Heading 2']
    for style_name in styles_to_change:
        try:
            style = doc.styles[style_name]
            style.font.name = '仿宋'
            style._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
        except KeyError:
            pass 
    
    h = doc.add_heading(f"【教学讲义】{topic_name}", 0)
    for run in h.runs:
        run.font.name = '仿宋'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    
    for i, slide in enumerate(ppt_data):
        h2 = doc.add_heading(f"环节 {i+1}：{slide.get('title', '无标题')}", level=2)
        for run in h2.runs:
            run.font.name = '仿宋'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
            
        for point in slide.get("content", []):
            p = doc.add_paragraph(point, style='List Bullet')
            for run in p.runs:
                run.font.name = '仿宋'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
        
        # 将 AI 生成的真实图片插入 Word
        if slide.get("image_bytes"):
            img_stream = io.BytesIO(slide["image_bytes"])
            doc.add_picture(img_stream, width=docx.shared.Inches(4.0))
            p = doc.add_paragraph("图：AI生成物理原理图")
            p.runs[0].font.italic = True
            for run in p.runs:
                run.font.name = '仿宋'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# ==========================================
# 3. 界面布局
# ==========================================
with st.sidebar:
    st.header("📂 本地资料库")
    uploaded_files = st.file_uploader("多选电脑教案 (按住 Command)：", type=['txt', 'docx'], accept_multiple_files=True)
    if uploaded_files:
        if st.button("📥 提取文件精华", use_container_width=True):
            with st.spinner("阅读中..."):
                all_content = [f"《{f.name}》\n{read_file(f)}" for f in uploaded_files]
                st.session_state.current_context = "\n\n".join(all_content)
                st.success("✅ 读取完毕！")
                
    st.markdown("---")
    topic = st.text_input("备课课题：", placeholder="例如：法拉第电磁感应定律")
    template_files = get_templates()
    selected_template = st.selectbox("PPT 模板：", template_files) if template_files else None
    selected_template_path = os.path.join("templates", selected_template) if selected_template else None

col_chat, col_studio = st.columns([6, 4], gap="large")

with col_chat:
    st.header("💬 智能教研控制台")
    chat_container = st.container(height=500)
    with chat_container:
        for msg in st.session_state.messages:
            st.chat_message(msg["role"]).write(msg["content"])

    st.markdown("⚡ **Agent 指令面板**")
    btn_col1, btn_col2 = st.columns(2)
    
    with btn_col1:
        if st.button("🕰️ 寻找硬核物理学史", use_container_width=True):
            if not topic: st.warning("先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"帮我找“{topic}”的真实学史和困境。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 翻阅历史文献中..."):
                            search_results = search_web(f"{topic} 物理学史 真实历史事件")
                            prompt = f"你是一位严谨的【科学史学者】。请为“{topic}”写一段硬核的引课，拒绝童话感，必须基于以下真实年份和困境：\n{search_results}"
                            response = client_brain.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.current_context += "\n\n" + reply 
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col2:
        if st.button("🪄 全能排版 (含AI生图)", type="primary", use_container_width=True):
            if not topic: st.warning("先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请为“{topic}”生成教案，并为重点页面配上插图指令。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        # 第 1 步：DeepSeek 写大纲
                        with st.spinner("🧠 大脑正在构思教案逻辑..."):
                            prompt = f"""
                            参考资料：{st.session_state.current_context}
                            请为课题“{topic}”生成大纲。
                            【输出格式】必须是纯 JSON 数组！如果这页需要配图，请在 'image_prompt' 中写出具体的画面描述（不需要则留空）：
                            [ {{"title": "标题", "content": ["要点1"], "image_prompt": "画面的具体描述"}} ]
                            """
                            response = client_brain.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.2)
                            result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                            ppt_data = json.loads(result_text)
                        
                        # 第 2 步：智谱 CogView 作图
                        with st.spinner("🎨 画师正在为您绘制物理插图..."):
                            for slide in ppt_data:
                                img_prompt = slide.get("image_prompt", "")
                                if img_prompt and img_prompt != "无":
                                    img_bytes = generate_physics_image(img_prompt)
                                    if img_bytes:
                                        slide["image_bytes"] = img_bytes
                            
                            st.session_state.ppt_data = ppt_data
                            reply = "✅ 图文教案已生成！插图和排版都在右侧 Studio，可以直接下载 PPT 和 Word 啦！"
                            st.write(reply)
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    if prompt := st.chat_input("和 Agent 聊聊教研..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.chat_message("user").write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    context_msg = [{"role": "system", "content": "你是资深物理老师。"}] + st.session_state.messages
                    response = client_brain.chat.completions.create(model="deepseek-chat", messages=context_msg, temperature=0.6)
                    reply = response.choices[0].message.content
                    st.write(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})

with col_studio:
    st.header("✨ Studio 成果区")
    if st.session_state.ppt_data is None:
        st.info("👈 在中栏点击【全能排版】后，带配图的终极教案将在这里诞生。")
    else:
        # 生成 PPT (把 AI 图片贴上去)
        if selected_template_path: prs = Presentation(selected_template_path)
        else: prs = Presentation() 

        for slide_data in st.session_state.ppt_data:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            if slide.shapes.title: slide.shapes.title.text = slide_data.get("title", "无标题")
            
            # 填入文字
            if len(slide.placeholders) > 1:
                tf = slide.placeholders[1].text_frame
                contents = slide_data.get("content", [""])
                if contents:
                    tf.text = contents[0]
                    for point in contents[1:]:
                        p = tf.add_paragraph()
                        p.text = point
                        p.level = 0
            
            # 贴入 AI 生成的真实图片 (贴在幻灯片右侧)
            if slide_data.get("image_bytes"):
                img_stream = io.BytesIO(slide_data["image_bytes"])
                try:
                    # 将图片放在距离左边缘 5 英寸，顶边缘 2 英寸的位置，宽度固定 4 英寸
                    slide.shapes.add_picture(img_stream, Inches(5), Inches(2), width=Inches(4))
                except Exception:
                    pass

        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        
        # 生成 Word
        word_stream = generate_word_document(st.session_state.ppt_data, topic if topic else "未命名课题")
        
        st.success("🎉 生成成功！请下载您的双重授课利器：")
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button("📊 下载带插图 PPT", data=ppt_stream, file_name=f"{topic}_带配图_AI备课.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
        with dl_col2:
            st.download_button("📝 下载带插图 仿宋Word", data=word_stream, file_name=f"{topic}_带配图_教案讲义.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="secondary", use_container_width=True)
        
        # 网页端预览
        st.markdown("### 👀 图文讲义预览")
        with st.container(height=500):
            for i, slide_data in enumerate(st.session_state.ppt_data):
                with st.container(border=True):
                    st.markdown(f"**环节 {i+1}：{slide_data.get('title', '无标题')}**")
                    for point in slide_data.get("content", []): st.markdown(f"- {point}")
                    # 在网页上直接把生成的图片展示出来
                    if slide_data.get("image_bytes"):
                        st.image(slide_data["image_bytes"], caption="AI 绘制的物理插图", width=300)