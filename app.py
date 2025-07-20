# app.py

import streamlit as st
from agent import create_excel_agent
from excel_utils import read_excel_file, preview_excel_sheets

st.set_page_config(page_title="🧠 Excel Sheets Agent", layout="wide")

st.title("📊 Excel Agent with LangChain")
st.markdown("Upload an Excel file and ask questions in natural language.")

# Step 1: Upload Excel File
uploaded_file = st.file_uploader("📤 Upload an Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        # Step 2: Read & Preview Sheets
        with st.spinner("🔍 Reading Excel..."):
            with open("temp_uploaded.xlsx", "wb") as f:
                f.write(uploaded_file.read())

            dfs = read_excel_file("temp_uploaded.xlsx")
            preview = preview_excel_sheets(dfs, rows=5)
            st.markdown("## 📑 Sheet Preview (First 5 Rows):")
            st.markdown(preview)

        # Step 3: Create Agent
        with st.spinner("🧠 Building LangChain agent..."):
            agent = create_excel_agent("temp_uploaded.xlsx", model_name="gpt-4")

        # Step 4: Query Box
        st.markdown("## 🤖 Ask a Question")
        user_input = st.text_input("Type your question and press Enter...")

        if user_input and agent:
            with st.spinner("⚙️ Thinking..."):
                try:
                    response = agent.invoke({"input": user_input})
                    st.success("✅ Done!")
                    st.markdown("### 📈 Answer")
                    st.markdown(response.get("output", "No output."))
                except Exception as e:
                    st.error(f"Error: {e}")

    except Exception as e:
        st.error(f"Failed to process file: {e}")
