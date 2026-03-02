import streamlit as st
from openai import OpenAI
import docx
from pptx import Presentation
import json
import io
import os

# 1. 页面基本设置
st.set_page_config(page_title="AI 物理备课助手", page_icon="⚛️", layout="wide")
st.title("⚛️ AI 物理备课助手 - 试题精讲版")
st.markdown("上传物理试题或解析，选择喜欢的模板，AI 将自动为你生成**「一题一页」**的精美幻灯片。")

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

# 3. 侧边栏
st.sidebar.header("🎨 幻灯片设置")
template_files = get_templates()

if not template_files:
    st.sidebar.warning("未检测到模板！请在 `templates` 文件夹中放入 .pptx 文件。当前将使用默认白底模板。")
    selected_template = None
else:
    selected_template = st.sidebar.selectbox("请选择一个 PPT 模板：", template_files)
    selected_template_path = os.path.join("templates", selected_template)

# 4. 主界面
st.markdown("### 第一步：上传试题或教案素材 (支持 .txt / .docx)")
uploaded_file = st.file_uploader("请上传你的文档：", type=['txt', 'docx'])

if uploaded_file is not None:
    file_content = read_file(uploaded_file)
    st.success("✅ 文件读取成功！")
    
    st.markdown("### 第二步：一键生成 PPT")
    if st.button("🪄 按「一题一页」生成 PPT"):
        with st.spinner('DeepSeek 正在拆解试题并套用模板，请稍候...'):
            try:
                # 给 AI 的提示词
                prompt = f"""
                请阅读以下物理教学素材（可能是试卷、题库或教案），将其转化为 PPT 大纲。
                核心要求：如果是试题，请务必做到【一道题单独占用一页幻灯片】。将题干、选项和解析放在同一页的要点中。
                
                请严格按照以下 JSON 数组格式输出，不要包含任何其他说明文字或 Markdown 标记：
                [
                    {{"title": "第1题 (或知识点标题)", "content": ["题目内容...", "A. ... B. ...", "【解析】..."]}},
                    {{"title": "第2题 (或知识点标题)", "content": ["题目内容...", "【解析】..."]}}
                ]
                
                教学素材内容如下：
                {file_content}
                """
                
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[
                        {"role": "system", "content": "你是一位严谨的高中物理特级教师。输出必须是纯 JSON 格式。"},
                        {"role": "user", "content": prompt},
                    ],
                    temperature=0.2 
                )
                
                # 解析 JSON
                result_text = response.choices[0].message.content
                result_text = result_text.replace("```json", "").replace("```", "").strip()
                ppt_data = json.loads(result_text)
                
                # 加载模板
                if selected_template:
                    prs = Presentation(selected_template_path)
                else:
                    prs = Presentation() 

                # 写入 PPT
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

                # 保存到内存
                ppt_stream = io.BytesIO()
                prs.save(ppt_stream)
                ppt_stream.seek(0)
                
                st.success("🎉 PPT 生成完毕！")

                # ==========================================
                # 新增功能：网页端内容卡片预览
                # ==========================================
                st.markdown("---")
                st.markdown("### 👀 幻灯片内容预览")
                
                # 使用两列排版，让预览看起来更紧凑
                cols = st.columns(2)
                for i, slide_data in enumerate(ppt_data):
                    col = cols[i % 2] # 左右列交替摆放
                    with col:
                        # 用带边框的容器模拟一张幻灯片
                        with st.container(border=True):
                            st.markdown(f"**第 {i+1} 页：{slide_data.get('title', '无标题')}**")
                            for point in slide_data.get("content", []):
                                st.markdown(f"- {point}")
                st.markdown("---")
                # ==========================================
                
                # 提供下载按钮
                st.download_button(
                    label="📥 对预览满意？点击下载生成的 PPTX 文件",
                    data=ppt_stream,
                    file_name="AI试题精讲.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except json.JSONDecodeError:
                st.error("❌ AI 返回的数据格式有误，请重新点击生成按钮试一下。")
            except Exception as e:
                st.error(f"❌ 发生错误：{e}")