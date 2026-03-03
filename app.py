# 核心改动 1：硬核物理学史与科学家引课 (拒绝低幼)
    with btn_col1:
        if st.button("🕰️ 找真实物理学史作引课", use_container_width=True):
            if not topic: st.warning("请先输入课题")
            else:
                st.session_state.messages.append({"role": "user", "content": f"请以真实客观的科学史视角，帮我挖掘“{topic}”的发现历程和真实历史细节，用来做硬核的课堂引课。"})
                with chat_container:
                    st.chat_message("user").write(st.session_state.messages[-1]["content"])
                    with st.chat_message("assistant"):
                        with st.spinner("🌍 正在检索深度的物理史料与真实文献..."):
                            # 修改搜索词：强制搜索传记、真实历史事件
                            search_results = search_web(f"{topic} 物理学史 真实历史事件 科学家传记 维基百科")
                            
                            # 终极 Prompt：压制 AI 的“童话感”，赋予“纪录片感”
                            prompt = f"""
                            你是一位严谨的【科学史学者】和资深的高中物理特级教师。
                            我需要你为课题“{topic}”准备一段具有【真实历史厚重感】的课堂引入。
                            
                            以下是我为你联网检索到的真实历史资料：
                            \n{search_results}\n
                            
                            要求：
                            1. 绝对拒绝低幼、虚构的童话式口吻（严禁出现“从前有个科学家”、“他灵机一动”这种轻浮表达）。
                            2. 必须基于真实的物理学史实。要具体到【真实的年份】、【当时的物理学界认知背景/瓶颈】以及【科学家面临的真实困境或实验误差】。
                            3. 语言风格要类似《典籍里的中国》或 BBC 科学纪录片的旁白，要有深度、有悬念，用科学的严谨来激发高中生的智力好奇心。
                            4. 如果检索资料中有科学家的真实名言或著作原话，请务必准确引用。
                            """
                            response = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}], temperature=0.3) # 温度调低至0.3，减少瞎编，增加客观事实
                            reply = response.choices[0].message.content
                            st.write(reply)
                            st.session_state.current_context += "\n\n" + reply 
                            st.session_state.messages.append({"role": "assistant", "content": reply})