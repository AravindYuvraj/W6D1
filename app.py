import streamlit as st
import pandas as pd
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain.agents import AgentExecutor, Tool, initialize_agent, AgentType
import re
import json
from dotenv import load_dotenv
import os
from thefuzz import process # New import for fuzzy matching

# --- Load environment variables from .env file ---
load_dotenv()

# --- Page Configuration ---
st.set_page_config(
    page_title="Intelligent Excel Agent",
    page_icon="🤖",
    layout="wide"
)

# --- Header ---
st.title("🤖 Intelligent Excel Agent (Street-Smart Edition)")
st.write("""
I can now understand fuzzy column names like "sales amount" instead of `revenue`!
Upload your Excel file, provide your Google API key, and ask me to analyze your data.
""")

# --- Helper Class: The "Excel Database" Connection ---
class ExcelDataFrame:
    def __init__(self, file_path):
        self.file_path = file_path
        try:
            self.dataframes = pd.read_excel(file_path, sheet_name=None)
            # Store original column names for fuzzy matching later
            self.original_columns = {sheet: df.columns.tolist() for sheet, df in self.dataframes.items()}
            
            for sheet_name, df in self.dataframes.items():
                df.columns = [self._clean_col_names(col) for col in df.columns]
                self.dataframes[sheet_name] = df
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {e}")

    def _clean_col_names(self, col_name):
        return re.sub(r'\W+', '_', str(col_name)).lower()

    def list_sheets(self):
        return list(self.dataframes.keys())

    def get_sheet_schema(self, sheet_name: str) -> str:
        sheet_name = sheet_name.strip().strip('"\'')
        if sheet_name not in self.dataframes:
            return f"Error: Sheet '{sheet_name}' not found."
        df = self.dataframes[sheet_name]
        if df.empty:
            return f"Sheet: '{sheet_name}' is empty."
        
        schema = f"Sheet: '{sheet_name}'\nCleaned Columns:\n"
        for col, dtype in df.dtypes.items():
            schema += f"- {col}: {dtype}\n"
        
        schema += "\nOriginal Columns:\n"
        for col in self.original_columns.get(sheet_name, []):
             schema += f"- {col}\n"
        return schema

    def get_dataframe(self, sheet_name: str) -> pd.DataFrame:
        sheet_name = sheet_name.strip().strip('"\'')
        if sheet_name not in self.dataframes:
            raise ValueError(f"Sheet '{sheet_name}' not found.")
        return self.dataframes[sheet_name]

# --- Core UI Components ---
st.sidebar.header("Setup")
google_api_key_input = st.sidebar.text_input("Enter your Google API Key:", type="password")
google_api_key = google_api_key_input or os.getenv("GOOGLE_API_KEY")

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
                llm = ChatGoogleGenerativeAI(model="gemini-2.0-flash", google_api_key=google_api_key, temperature=0)

                def analyze_sheet(tool_input: str) -> str:
                    try:
                        if "```json" in tool_input:
                            tool_input = tool_input.replace("```json\n", "").replace("\n```", "")
                        
                        input_data = json.loads(tool_input)
                        sheet_name = input_data["sheet_name"]
                        query = input_data["query"]

                        st.write(f"🔍 Analyzing sheet: **{sheet_name}** with query: \"{query}\"")
                        
                        df = excel_db.get_dataframe(sheet_name)
                        if df.empty:
                            return f"The sheet '{sheet_name}' is empty and cannot be analyzed."

                        # --- FIX: Make the specialist agent's prompt extremely direct ---
                        specialist_prefix = """
                        You are a Python code executor.
                        You are working with a pandas dataframe in Python named `df`.
                        Your task is to write and execute a single line of Python code to answer the user's query.
                        You MUST output only the raw result of the code execution.
                        Do not add any conversational text, explanations, or thoughts.
                        Just return the direct output from the python_repl_ast tool.
                        """

                        specialist_agent_executor = create_pandas_dataframe_agent(
                            llm,
                            df,
                            prefix=specialist_prefix,
                            verbose=True,
                            agent_executor_kwargs={"handle_parsing_errors": True},
                            allow_dangerous_code=True
                        )
                        
                        result = specialist_agent_executor.invoke({"input": query})
                        # The specialist now returns the raw result, so the manager must add the conversational text.
                        return f"The result of the analysis is: {result['output']}"

                    except Exception as e:
                        return f"An error occurred during analysis: {e}"

                def list_sheets_tool(ignored_input=""):
                    return excel_db.list_sheets()

                def get_column_name_match(tool_input: str) -> str:
                    try:
                        input_data = json.loads(tool_input)
                        sheet_name = input_data["sheet_name"]
                        column_name = input_data["column_name"]

                        actual_columns = excel_db.original_columns.get(sheet_name)
                        if not actual_columns:
                            return f"Error: Could not find columns for sheet '{sheet_name}'."

                        best_match, score = process.extractOne(column_name, actual_columns)
                        if score > 80: # Confidence threshold
                            return excel_db._clean_col_names(best_match)
                        else:
                            return f"Error: Could not find a confident match for column '{column_name}' in sheet '{sheet_name}'. Best guess was '{best_match}'."
                    except Exception as e:
                        return f"An error occurred during column matching: {e}"


                tools = [
                    Tool(
                        name="list_sheets",
                        func=list_sheets_tool,
                        description="Use this tool to get a list of all sheet names in the Excel file. Call it without any arguments.",
                    ),
                    Tool(
                        name="get_sheet_schema",
                        func=excel_db.get_sheet_schema,
                        description="Use this tool to get the schema (column names and data types) of a specific sheet. Pass the sheet name as an argument.",
                    ),
                    Tool(
                        name="get_column_name_match",
                        func=get_column_name_match,
                        description="Use this tool to find the correct, cleaned column name that matches a user's fuzzy column name. Input must be a JSON with 'sheet_name' and 'column_name'."
                    ),
                    Tool(
                        name="analyze_sheet",
                        func=analyze_sheet,
                        description="Use this tool to perform detailed analysis on a specific sheet, AFTER you have confirmed all column names are correct. Input must be a JSON with 'sheet_name' and 'query'."
                    )
                ]
                
                system_message = """
                You are a helpful and friendly "Manager" agent for an Excel file.
                Your primary goal is to understand a user's query and delegate tasks to the correct tools.

                **Your Workflow:**

                1.  **Understand the Goal:** Analyze the user's question to determine what they want to achieve.

                2.  **Identify Sheets and Columns:**
                    * If you don't know which sheet to use, call `list_sheets`.
                    * If the user's query mentions column names (e.g., "sales amount", "Customer ID"), you **MUST** translate them to the correct column names before analyzing the data.

                3.  **Translate Column Names (if necessary):**
                    * For each fuzzy column name in the user's query, use the `get_column_name_match` tool.
                    * **Example:** If the user asks "What is the total sales amount?", you must first call `get_column_name_match`. Your Action Input for this tool should be a JSON object with the `sheet_name` key set to "Sales" and the `column_name` key set to "sales amount". The tool will return the correct name, e.g., `revenue`.

                4.  **Analyze the Data:**
                    * Once you have the correct sheet name and all column names have been translated, construct a new query using the *correct* column names.
                    * Call the `analyze_sheet` tool with the corrected query.
                    * **Example:** After translating "sales amount" to `revenue`, you call `analyze_sheet`. Your Action Input for this tool should be a JSON object with the `sheet_name` key set to "Sales" and the `query` key set to "What is the total revenue?".

                **Crucial Rule:** Never call `analyze_sheet` with a user's raw query if it contains column names. Always translate column names first using `get_column_name_match`.
                """
                
                agent_kwargs = {
                    "prefix": system_message,
                }

                manager_agent_executor = initialize_agent(
                    tools,
                    llm,
                    agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION,
                    verbose=True,
                    handle_parsing_errors=True,
                    agent_kwargs=agent_kwargs
                )
                
                st.session_state.agent_executor = manager_agent_executor
                st.sidebar.success("Street-Smart Agent Initialized!")
                if not st.session_state.messages:
                    st.session_state.messages.append({"role": "assistant", "content": "Hello! I can now understand fuzzy column names. What would you like to know?"})

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
