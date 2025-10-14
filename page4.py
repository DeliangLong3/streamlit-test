import streamlit as st

with st.sidebar:
    name=st.text_input("请输入你的名字")
    if name:
        st.write(f"你好,{name}，谢谢你参与测试")


column1,column2,column3=st.columns([1,1,1])
with column1:
    password=st.text_input("你能打开这个网页吗？：")
    if password:
        st.write(f"谢谢你的尝试，请继续回答下一个问题。")

with column2:
    paragraph=st.text_area("你觉得这个网页看起来简单吗？")
    if paragraph:
        st.write(f"看来你认为这个网页挺{paragraph}嘛！😄")

with column3:
    score=st.slider("请给这个网页打个分数吧：",value=None,max_value=100,step=10)
    if score:
        st.write(f"你打的分数是：{score}分，谢谢你的评估和打分。")

st.divider()

checked=st.checkbox("已完成本网页测评。")
if checked:
    st.write("感谢你的测评")

st.divider()

submitted=st.button("提交")
if submitted:
    st.write("提交成功！")