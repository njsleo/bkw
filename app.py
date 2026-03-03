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
    st.session_state.messages = [{"role": "assistant", "content": "老师好！现在您可以：\n1. 在左侧自由勾选想要分析的文档。\n2. 在下方对我说：**“把选中文档的前三题导出为 Word”** 或 **“把这部分做成 PPT”**，我会100%忠于原文档，并只为您生成需要的格式！"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "current_context" not in st.session_state:
    st.session_state.current_context = ""
if "output_type" not in st.session_state:
    st.session_state.output_type = "ppt" # 默认输出格式

# ==========================================
# 2. 核心功能引擎
# ==========================================
def read_file(file_obj):
    try:
        if file_obj.name.endswith('.txt'):
            content_bytes = file_obj.getvalue()
            try: return content_bytes.decode("utf-8")
            except UnicodeDecodeError:
                try: return content_bytes.decode("gbk")
                except Exception: return content_bytes.decode("utf-8", errors="ignore")
        elif file_obj.name.endswith('.docx'):
            return "\n".join([para.text for para in docx.Document(file_obj).paragraphs])
    except Exception as e:
        return f"[读取失败: {e}]"
    return ""

def get_templates():
    if not os.path.exists("templates"): return []
    return [f for f in os.listdir("templates") if f.endswith(".pptx")]

def search_web(query):
    try:
        results = DDGS().text(query, max_results=3)
        return "\n".join([f"- 【{r['title']}】: {r['body']}" for r in results])
    except Exception: return "搜索失败"

def generate_physics_image(image_prompt):
    try:
        response = client_paint.images.generations(
            model="cogview-3-plus", 
            prompt=f"一张专业的高中物理教学幻灯片配图。画面内容：{image_prompt}。要求：纯白背景，绝不要出现乱码英文字母或公式。",
        )
        img_data = requests.get(response.data[0].url).content
        return img_data
    except Exception: return None

def generate_word_document(ppt_data, topic_name):
    doc = docx.Document()
    styles_to_change = ['Normal', 'List Bullet', 'Title', 'Heading 1', 'Heading 2']
    for style_name in styles_to_change:
        try:
            style = doc.styles[style_name]
            style.font.name = '仿宋'
            style._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
        except KeyError: pass 
    
    h = doc.add_heading(f"【教研输出】{topic_name}", 0)
    for run in h.runs: run.font.name = '仿宋'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    
    for i, slide in enumerate(ppt_data):
        h2 = doc.add_heading(f"{slide.get('title', '无标题')}", level=2)
        for run in h2.runs: run.font.name = '仿宋'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
            
        for point in slide.get("content", []):
            p = doc.add_paragraph(point, style='List Bullet')
            for run in p.runs: run.font.name = '仿宋'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
        
        image_type = slide.get("image_type", "none")
        if image_type == "creative" and slide.get("image_bytes"):
            doc.add_picture(io.BytesIO(slide["image_bytes"]), width=docx.shared.Inches(4.0))
            p = doc.add_paragraph("图：AI生成创意配图"); p.runs[0].font.italic = True
        elif image_type == "schematic":
            p = doc.add_paragraph(f"✂️ [此处预留空位，请手动截取原卷题目图片粘贴]")
            p.runs[0].font.color.rgb = docx.shared.RGBColor(0, 112, 192)
            
        for p in doc.paragraphs[-2:]: 
            for run in p.runs: run.font.name = '仿宋'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# ==========================================
# 3. 界面布局
# ==========================================
with st.sidebar:
    st.header("📂 本地资料库")
    uploaded_files = st.file_uploader("批量上传教案/习题：", type=['txt', 'docx'], accept_multiple_files=True)
    
    if uploaded_files:
        # 核心修复 1：新增文档自由勾选功能
        file_names = [f.name for f in uploaded_files]
        selected_names = st.multiselect("☑️ 请选择本次要让 AI 分析的文档：", options=file_names, default=file_names)
        
        if st.button("📥 仅读取选中的文档", use_container_width=True):
            if not selected_names:
                st.warning("请至少勾选一个文档！")
            else:
                with st.spinner("正在精准解析勾选的文档..."):
                    selected_files = [f for f in uploaded_files if f.name in selected_names]
                    all_content = [f"《{f.name}》\n{read_file(f)}" for f in selected_files]
                    if all_content:
                        st.session_state.current_context = "\n\n".join(all_content)
                        st.success(f"✅ 成功读取 {len(selected_files)} 份文档！AI 目前只会基于这些文档思考。")
                
    st.markdown("---")
    topic = st.text_input("备课课题 (仅作大标题)：", placeholder="例如：法拉第电磁感应定律")
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

    # ==========================================
    # 核心修复 2 & 3：紧箍咒 Prompt 与 智能格式分发
    # ==========================================
    if prompt := st.chat_input("输入指令：把前三道题导出为Word... / 提炼考点做成PPT..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with chat_container:
            st.chat_message("user").write(prompt)
            with st.chat_message("assistant"):
                
                lower_prompt = prompt.lower()
                is_generation_intent = any(word in lower_prompt for word in ["ppt", "幻灯片", "排版", "做成", "word", "教案", "讲义", "导出为", "生成"])
                
                if is_generation_intent:
                    # 判断用户需要什么格式
                    if any(w in lower_prompt for w in ["word", "教案", "讲义", "文档"]):
                        st.session_state.output_type = "word"
                        format_str = "Word 讲义"
                    else:
                        st.session_state.output_type = "ppt"
                        format_str = "PPT 幻灯片"

                    with st.spinner(f"🧠 收到指令，正在严格基于文档生成 {format_str}..."):
                        
                        # 核心“紧箍咒” Prompt：严禁脱离文档
                        ppt_prompt = f"""
                        【最高指令】：你必须绝对忠实于以下我提供的本地提取资料，绝不允许脱离资料自己发散或编造！
                        
                        以下是供你分析的本地提取资料：
                        \n{st.session_state.current_context}\n
                        
                        用户的新指令是：“{prompt}”
                        
                        请你严格按照用户的指令（如：只提取前三道题），从上面的资料中提取对应内容。
                        排版建议：如果提取的是题目，title设为“题目X”，content中原封不动保留题干和选项。
                        
                        配图判断：
                        - 习题图/电路图/受力图 -> "schematic" (image_prompt 留空，系统会预留手工截图位)
                        - 宏观场景/历史人物 -> "creative" (并在 image_prompt 写英文提示词)
                        - 不需要图 -> "none"
                        
                        【输出格式】必须是纯 JSON 数组：
                        [ {{"title": "题目1", "content": ["题干...", "A.选项"], "image_type": "schematic", "image_prompt": ""}} ]
                        """
                        try:
                            response = client_brain.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": ppt_prompt}], temperature=0.1) # 温度降到 0.1，极其严谨，不瞎编
                            result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
                            ppt_data = json.loads(result_text)
                            
                            # 只有选了需要创意图才去生图
                            for slide in ppt_data:
                                if slide.get("image_type") == "creative" and slide.get("image_prompt"):
                                    img_bytes = generate_physics_image(slide["image_prompt"])
                                    if img_bytes: slide["image_bytes"] = img_bytes
                            
                            st.session_state.ppt_data = ppt_data
                            reply = f"✅ 没问题！我已经严格按照您提取的文档内容，生成了专属的 **{format_str}** 并发送至右侧 Studio，请查收！"
                            st.write(reply)
                            st.session_state.messages.append({"role": "assistant", "content": reply})
                            st.rerun() # 强制刷新渲染右侧
                            
                        except Exception as e:
                            err = f"❌ 提取生成失败，请检查文档或指令。（报错：{e}）"
                            st.error(err)
                            st.session_state.messages.append({"role": "assistant", "content": err})
                            
                else:
                    # 普通提问
                    with st.spinner("思考中..."):
                        sys_prompt = "你是资深物理老师。"
                        if st.session_state.current_context:
                            sys_prompt += f"\n\n请严格基于以下资料回答：\n{st.session_state.current_context}"
                        
                        context_msg = [{"role": "system", "content": sys_prompt}] + st.session_state.messages
                        response = client_brain.chat.completions.create(model="deepseek-chat", messages=context_msg, temperature=0.6)
                        reply = response.choices[0].message.content
                        st.write(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})

with col_studio:
    st.header("✨ Studio 成果区")
    if st.session_state.ppt_data is None:
        st.info("👈 左侧勾选文档并提取后，在聊天框发送导出指令（如：把前三题导出为Word）。")
    else:
        # 核心修复 4：根据指令按需渲染，绝不多浪费一秒钟
        is_word_mode = (st.session_state.output_type == "word")
        
        st.success(f"🎉 您的 {'Word 讲义' if is_word_mode else 'PPT 幻灯片'} 生成成功！")
        
        if is_word_mode:
            # 只渲染 Word
            with st.spinner("正在生成排版精美的 Word 文件..."):
                word_stream = generate_word_document(st.session_state.ppt_data, topic if topic else "文档提取")
                st.download_button("📝 立即下载 仿宋 Word", data=word_stream, file_name=f"{topic}_教案输出.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
        else:
            # 只渲染 PPT
            with st.spinner("正在生成带配图的 PPT 文件..."):
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
                    
                    image_type = slide_data.get("image_type", "none")
                    if image_type == "creative" and slide_data.get("image_bytes"):
                        try: slide.shapes.add_picture(io.BytesIO(slide_data["image_bytes"]), Inches(5), Inches(2), width=Inches(4))
                        except Exception: pass
                    elif image_type == "schematic":
                        try:
                            txBox = slide.shapes.add_textbox(Inches(5), Inches(2), Inches(4), Inches(3))
                            p = txBox.text_frame.paragraphs[0]
                            p.text = "✂️ 【截图预留位】\n请在此处使用快捷键\n粘贴原卷中的习题配图"
                            p.font.color.rgb = RGBColor(0, 112, 192)
                            p.font.size = Pt(20)
                        except Exception: pass

                ppt_stream = io.BytesIO()
                prs.save(ppt_stream)
                ppt_stream.seek(0)
                st.download_button("📊 立即下载 智能双轨 PPT", data=ppt_stream, file_name=f"{topic}_大纲提取.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
        
        # 共同的预览视图
        st.markdown(f"### 👀 内容预览 ({'Word模式' if is_word_mode else 'PPT模式'})")
        with st.container(height=500):
            for i, slide_data in enumerate(st.session_state.ppt_data):
                with st.container(border=True):
                    st.markdown(f"**{slide_data.get('title', '内容')}**")
                    for point in slide_data.get("content", []): st.markdown(f"{point}")
                    
                    if slide_data.get("image_type") == "creative" and slide_data.get("image_bytes"):
                        st.image(slide_data["image_bytes"], caption="AI 绘制的创意配图", width=300)
                    elif slide_data.get("image_type") == "schematic":
                        st.info("✂️ Agent 已为本页预留【原题截图空位】")