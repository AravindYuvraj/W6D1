import streamlit as st
import pandas as pd
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain.tools import Tool
import re

# --- Page Configuration ---
st.set_page_config(
    page_title="Intelligent Excel Agent",
    page_icon="🤖",
    layout="wide"
)

# --- Header ---
st.title("🤖 Intelligent Excel Agent (Gemini Edition)")
st.write("""
Welcome! I am an intelligent agent designed to help you analyze and understand your Excel data.
Upload your Excel file, provide your Google API key, and then ask me questions about it!
""")

# --- Helper Class: The "Excel Database" Connection ---
class ExcelDataFrame:
    def __init__(self, file_path):
        self.file_path = file_path
        try:
            self.dataframes = pd.read_excel(file_path, sheet_name=None)
            for sheet_name, df in self.dataframes.items():
                df.columns = [self._clean_col_names(col) for col in df.columns]
                self.dataframes[sheet_name] = df
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {e}")

    def _clean_col_names(self, col_name):
        """Cleans column names to be valid Python identifiers."""
        return re.sub(r'\W+', '_', col_name).lower()

    def list_sheets(self):
        """Returns a list of sheet names."""
        return list(self.dataframes.keys())

    def get_sheet_schema(self, sheet_name: str) -> str:
        """
        Returns a description of the specified sheet's schema (columns and data types).
        """
        if sheet_name not in self.dataframes:
            return f"Error: Sheet '{sheet_name}' not found."
        df = self.dataframes[sheet_name]
        if df.empty:
            return f"Sheet: '{sheet_name}' is empty."
            
        schema = "Sheet: '{}'\nColumns:\n".format(sheet_name)
        for col, dtype in df.dtypes.items():
            schema += f"- {col}: {dtype}\n"
        return schema

    def get_dataframe(self, sheet_name: str) -> pd.DataFrame:
        """Returns the pandas DataFrame for the specified sheet."""
        if sheet_name not in self.dataframes:
            raise ValueError(f"Sheet '{sheet_name}' not found.")
        return self.dataframes[sheet_name]

# --- Core UI Components ---
st.sidebar.header("Setup")
google_api_key = st.sidebar.text_input("Enter your Google API Key:", type="password")
uploaded_file = st.sidebar.file_uploader(
    "Choose an Excel file (.xlsx)",
    type="xlsx"
)

# --- Main App Logic & Agent Initialization ---
if 'agent_executor' not in st.session_state:
    st.session_state.agent_executor = None
    st.session_state.messages = []

if uploaded_file and google_api_key:
    if st.session_state.agent_executor is None:
        with st.spinner("Initializing Agent... Please wait."):
            try:
                excel_db = ExcelDataFrame(uploaded_file)

                tools = [
                    Tool(
                        name="list_sheets",
                        func=lambda x: excel_db.list_sheets(),
                        description="Use this tool to get a list of all sheet names in the Excel file.",
                    ),
                    Tool(
                        name="get_sheet_schema",
                        func=excel_db.get_sheet_schema,
                        description="Use this tool to get the schema (column names and data types) of a specific sheet. Pass the sheet name as the argument.",
                    )
                ]

                # --- The Agent's Brain (LLM + Prompt) ---
                # Switched from OpenAI to Google Gemini
                llm = ChatGoogleGenerativeAI(model="gemini-2.0-flash", google_api_key=google_api_key, temperature=0)

                system_message = """
                You are a helpful and friendly agent designed to interact with an Excel file.
                Your goal is to help the user understand their data.
                You have tools to list the sheets in the file and to see the schema of each sheet.
                - ALWAYS start by using the `list_sheets` tool to see what's available.
                - Do not assume column names or sheet names. Use your tools to find them.
                - If the user asks about a specific sheet, use the `get_sheet_schema` tool to understand its structure before answering.
                - If you don't know the answer, say that you don't have the tools to answer that question yet.
                - The user is interacting with you through a chat interface, so provide conversational and clear answers.
                """
                
                agent = create_pandas_dataframe_agent(
                    llm,
                    pd.DataFrame(), # Dummy dataframe
                    prefix=system_message,
                    agent_executor_kwargs={"handle_parsing_errors": True},
                    extra_tools=tools,
                    verbose=True,
                    allow_dangerous_code=True # FIX: Explicitly allow code execution
                )
                
                st.session_state.agent_executor = agent
                st.sidebar.success("Agent Initialized!")
                if not st.session_state.messages:
                    st.session_state.messages.append({"role": "assistant", "content": "Hello! I'm ready to help you with your Excel file. What would you like to know?"})

            except Exception as e:
                st.sidebar.error(f"Initialization failed: {e}")
                st.session_state.agent_executor = None

# --- Chat Interface ---
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.write(message["content"])

if prompt := st.chat_input("Ask a question about your Excel file..."):
    if not google_api_key:
        st.info("Please add your Google API key to continue.")
        st.stop()
    if not uploaded_file:
        st.info("Please upload an Excel file to continue.")
        st.stop()
    if st.session_state.agent_executor is None:
        st.info("Agent not initialized. Please ensure API key and file are provided.")
        st.stop()

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.write(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Thinking..."):
            try:
                response = st.session_state.agent_executor.invoke({"input": prompt})
                st.write(response["output"])
                st.session_state.messages.append({"role": "assistant", "content": response["output"]})
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.session_state.messages.append({"role": "assistant", "content": f"Sorry, I ran into an error: {e}"})