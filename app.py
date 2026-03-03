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
    st.session_state.messages = [{"role": "assistant", "content": "老师好！我已经连接了您的**专属云盘**，并切换至**物理学史引课模式**。您可以随时让我从云盘里提取资料，或者帮您构思引课故事。"}]
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None
if "current_context" not in st.session_state:
    st.session_state.current_context = ""

# ==========================================
# 2. 辅助功能函数
# ==========================================
def read_file(file_path_or_file_obj):
    """支持读取上传的文件对象，也支持读取本地云盘的路径"""
    try:
        if hasattr(file_path_or_file_obj, 'name'): # 如果是网页上传的文件
            name = file_path_or_file_obj.name
            if name.endswith('.txt'):
                return file_path_or_file_obj.getvalue().decode("utf-8")
            elif name.endswith('.docx'):
                return "\n".join([para.text for para in docx.Document(file_path_or_file_obj).paragraphs])
        else: # 如果是云盘里的本地文件
            if file_path_or_file_obj.endswith('.txt'):
                with open(file_path_or_file_obj, 'r', encoding='utf-8') as f:
                    return f.read()
            elif file_path_or_file_obj.endswith('.docx'):
                return "\n".join([para.text for para in docx.Document(file_path_or_file_obj).paragraphs])
    except Exception as e:
        return f"读取 {file_path_or_file_obj} 失败: {e}"
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

# 读取云盘中所有资料
def read_cloud_drive():
    cloud_dir = "my_cloud_drive"
    if not os.path.exists(cloud_dir):
        return "您的云盘目前是空的。请在 VS Code 中创建 `my_cloud_drive` 文件夹并放入资料。"
    
    files = [f for f in os.listdir(cloud_dir) if f.endswith(('.txt', '.docx'))]
    if not files:
        return "云盘中未发现 .txt 或 .docx 文件。"
    
    all_content = []
    for f in files:
        file_path = os.path.join(cloud_dir, f)
        all_content.append(f"--- 来自云盘文件《{f}》 --- \n" + read_file(file_path))
    
    return "\n\n".join(all_content)

# ==========================================
# 3. 界面布局：左、中、右
# ==========================================

with st.sidebar:
    st.header("☁️ 我的云盘资料库")
    st.markdown("只要将教案放入 `my_cloud_drive` 文件夹并同步，这里就能自动读取。")
    
    if st.button("🔄 扫描并加载云盘资料", use_container_width=True):
        with st.spinner("正在读取云盘文件..."):
            cloud_content = read_cloud_drive()
            st.session_state.current_context = cloud_content
            st.success("云盘资料已成功加载到 AI 记忆中！")
            with st.expander("查看云盘加载内容"):
                st.write(cloud_content[:500] + "......")

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
    
    # 核心改动 1：物理学史与科学家引课
    with btn_col1:
        if st.button("🕰️ 找物理学史作引课", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"帮我搜索“{topic}”的发现历程和相关科学家的故事，用来做课堂引课。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 正在检索物理学史与科学家轶事..."):
                            # 修改搜索词，精准打击学史和出处
                            search_results = search_web(f"{topic} 物理学史 发现过程 科学家故事 起源")
                            
                            prompt = f"""
                            你是一位极具魅力的物理老师。用户想找关于“{topic}”的引课素材。我为你联网搜索到了以下历史资料：
                            \n{search_results}\n
                            请你根据这些资料，以该知识点的【出处、发展历程和核心科学家】为切入点，用讲故事的口吻写一段引人入胜的课堂开场白。
                            """
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.7)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.current_context += "\n\n" + reply 
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    # 核心改动 2：直接从云盘里找资料
    with btn_col2:
        if st.button("☁️ 从我的云盘提取", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            elif not st.session_state.current_context: st.warning("请先在左侧点击【扫描并加载云盘资料】")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请结合我云盘里的资料，提取关于“{topic}”的备课核心要点。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🧠 正在云盘中筛选相关知识..."):
                            prompt = f"这是我云盘里的所有教案资料：\n{st.session_state.current_context}\n\n请你从中找出与“{topic}”最相关的内容，并整理成备课要点。"
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3)
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.messages.append({"role": "assistant", "content": reply})

    with btn_col3:
        if st.button("🪄 一键做成 PPT", type="primary", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请根据我们刚聊的内容（包括学史引课和云盘资料），为“{topic}”生成PPT。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🎨 排版中..."):
                            # 结合聊天记录和云盘资料生成PPT
                            chat_history = "\n".join([m["content"] for m in st.session_state.messages[-3:]])
                            prompt = f"""
                            参考资料：{chat_history}
                            请结合参考资料，为课题“{topic}”生成 PPT 大纲。确保把【物理学史/科学家故事】放在PPT的前几页作为引入。
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