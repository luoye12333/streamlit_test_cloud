import streamlit as st

st.title("你好，Streamlit Cloud！")
name = st.text_input("请输入你的名字：")
if name:
    st.success(f"欢迎，{name}！")