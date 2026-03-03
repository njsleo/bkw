import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
import json
import io
import os

# 1. 页面基本设置
st.set_page_config(page_title="AI 物理备课助手", page_icon="⚛️", layout="wide")
st.title("⚛️ AI 物理备课助手 - 旗舰版")
st.markdown("选择你的备课方式：让 AI 帮你**提炼文档**，或者让 AI 为你**头脑风暴**。")

# 2. 读取 API Key
try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except KeyError:
    st.error("⚠️ 未找到 API Key！请检查 Streamlit 后台的 Secrets 配置。")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

# --- 辅助函数 ---
def read_file(uploaded_file):
    if uploaded_file.name.endswith('.txt'):
        return uploaded_file.getvalue().decode("utf-8")
    elif uploaded_file.name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        return "\n".join([para.text for para in doc.paragraphs])
    return None

def get_templates():
    if not os.path.exists("templates"):
        return []
    return [f for f in os.listdir("templates") if f.endswith(".pptx")]

# --- 核心：生成 PPT 并提供下载的函数 ---
def create_ppt_and_preview(ppt_data, template_path):
    # 1. 网页端卡片预览
    st.markdown("---")
    st.markdown("### 👀 幻灯片内容预览")
    cols = st.columns(2)
    for i, slide_data in enumerate(ppt_data):
        col = cols[i % 2]
        with col:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 页：{slide_data.get('title', '无标题')}**")
                for point in slide_data.get("content", []):
                    st.markdown(f"- {point}")
    st.markdown("---")

    # 2. 生成真实的 PPT 文件
    if template_path:
        prs = Presentation(template_path)
    else:
        prs = Presentation() 

    for slide_data in ppt_data:
        slide_layout = prs.slide_layouts[1] 
        slide = prs.slides.add_slide(slide_layout)
        if slide.shapes.title:
            slide.shapes.title.text = slide_data.get("title", "无标题")
        if len(slide.placeholders) > 1:
            body_shape = slide.placeholders[1]
            tf = body_shape.text_frame
            contents = slide_data.get("content", [""])
            if contents:
                tf.text = contents[0]
                for point in contents[1:]:
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0

    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    
    # 3. 提供下载
    st.download_button(
        label="📥 方案设计得不错！点击下载 PPT",
        data=ppt_stream,
        file_name="AI物理备课.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# 3. 侧边栏：模板选择
st.sidebar.header("🎨 幻灯片设置")
template_files = get_templates()

if not template_files:
    st.sidebar.warning("未检测到模板！请在 `templates` 文件夹中放入 .pptx 文件。")
    selected_template = None
    selected_template_path = None
else:
    selected_template = st.sidebar.selectbox("请选择一个 PPT 模板：", template_files)
    selected_template_path = os.path.join("templates", selected_template)

# ==========================================
# 4. 主界面：双模式选项卡 (Tabs)
# ==========================================
tab1, tab2 = st.tabs(["💡 从主题灵感生成 (AI 帮你设计)", "📄 从本地文档提取 (AI 帮你排版)"])

# ----------------- 模式 A：从主题生成 -----------------
with tab1:
    st.markdown("### 告诉 AI 你想讲什么课？")
    topic = st.text_input("请输入教学主题（例如：楞次定律、动量守恒定律的应用、平抛运动）：", placeholder="例如：带电粒子在匀强磁场中的运动")
    
    custom_req_topic = st.text_area(
        "对教学设计的特殊要求 (选填)：", 
        height=100,
        key="req_topic",
        placeholder="例如：'引入部分要结合生活中有趣的例子'、'重点放在受力分析上'、'最后加一道高考真题做结尾'"
    )
    
    if st.button("🚀 让 AI 开始备课设计", key="btn_topic"):
        if not topic.strip():
            st.warning("请先输入一个教学主题哦！")
        else:
            with st.spinner(f'DeepSeek 正在为你精心设计【{topic}】的教学大纲，请稍候...'):
                try:
                    instruction_text = f"\n【⚠️ 用户的特殊要求】：\n{custom_req_topic}\n" if custom_req_topic.strip() else ""
                    prompt = f"""
                    你是一位资深的高中物理特级教师。现在请你围绕教学主题“{topic}”，**凭空设计**一份逻辑严密、循序渐进的 PPT 教学大纲。
                    包含环节建议：课堂引入、核心概念解析、公式推导、例题精讲、课堂小结。
                    {instruction_text}
                    
                    无论设计什么内容，【输出格式】必须绝对遵守以下 JSON 数组格式，绝对不要包含任何其他说明文字或 Markdown 标记：
                    [
                        {{"title": "幻灯片标题", "content": ["要点1", "要点2"]}},
                        {{"title": "幻灯片标题", "content": ["要点1", "要点2"]}}
                    ]
                    """
                    
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": "你是一位优秀的物理老师。输出必须是纯 JSON 格式。"},
                            {"role": "user", "content": prompt},
                        ],
                        temperature=0.6 # 稍微提高一点温度，让 AI 更有创造力
                    )
                    
                    result_text = response.choices[0].message.content
                    result_text = result_text.replace("```json", "").replace("```", "").strip()
                    ppt_data = json.loads(result_text)
                    st.success("🎉 AI 教学设计完成！请审阅下方大纲。")
                    
                    # 调用统一的预览和下载函数
                    create_ppt_and_preview(ppt_data, selected_template_path)
                    
                except json.JSONDecodeError:
                    st.error("❌ AI 返回格式有误，请重试。")
                except Exception as e:
                    st.error(f"❌ 发生错误：{e}")

# ----------------- 模式 B：从文档提取 -----------------
with tab2:
    st.markdown("### 上传试题或教案素材")
    uploaded_file = st.file_uploader("支持 .txt 或 .docx 格式：", type=['txt', 'docx'])
    
    custom_req_doc = st.text_area(
        "对 AI 的特殊要求 (选填)：", 
        height=100,
        key="req_doc",
        placeholder="例如：'如果是试卷，请做到一题一页'、'只要题干不需要解析'"
    )

    if uploaded_file is not None:
        file_content = read_file(uploaded_file)
        st.success("✅ 文件读取成功！")
        
        if st.button("🪄 按文档内容生成 PPT", key="btn_doc"):
            with st.spinner('DeepSeek 正在拆解文档，请稍候...'):
                try:
                    instruction_text = f"\n【⚠️ 用户的特殊要求】：\n{custom_req_doc}\n" if custom_req_doc.strip() else ""
                    prompt = f"""
                    请阅读以下物理教学素材，将其转化为 PPT 大纲。
                    {instruction_text}
                    
                    【输出格式】必须绝对遵守以下 JSON 数组格式，不要包含任何说明文字：
                    [
                        {{"title": "幻灯片标题", "content": ["要点1", "要点2"]}}
                    ]
                    
                    素材内容：
                    {file_content}
                    """
                    
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": "你是一位严谨的物理老师。输出必须是纯 JSON 格式。"},
                            {"role": "user", "content": prompt},
                        ],
                        temperature=0.2 
                    )
                    
                    result_text = response.choices[0].message.content
                    result_text = result_text.replace("```json", "").replace("```", "").strip()
                    ppt_data = json.loads(result_text)
                    st.success("🎉 提取完毕！请审阅下方幻灯片内容。")
                    
                    # 调用统一的预览和下载函数
                    create_ppt_and_preview(ppt_data, selected_template_path)
                    
                except json.JSONDecodeError:
                    st.error("❌ AI 返回格式有误，请重试。")
                except Exception as e:
                    st.error(f"❌ 发生错误：{e}")