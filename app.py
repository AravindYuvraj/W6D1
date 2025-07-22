import streamlit as st
import pandas as pd
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain.agents import AgentExecutor, Tool, initialize_agent, AgentType
import re
import json
from dotenv import load_dotenv
import os

# --- Load environment variables from .env file ---
load_dotenv()

# --- Page Configuration ---
st.set_page_config(
    page_title="Intelligent Excel Agent",
    page_icon="🤖",
    layout="wide"
)

# --- Header ---
st.title("🤖 Intelligent Excel Agent (Final Analyst Edition)")
st.write("""
This version uses a robust agent creation method to solve previous errors.
Upload your Excel file, provide your Google API key, and ask me to analyze your data.
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
        
        schema = f"Sheet: '{sheet_name}'\nColumns:\n"
        for col, dtype in df.dtypes.items():
            schema += f"- {col}: {dtype}\n"
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
                llm = ChatGoogleGenerativeAI(model="gemini-1.5-flash", google_api_key=google_api_key, temperature=0)

                def analyze_sheet(tool_input: str) -> str:
                    try:
                        if "```json" in tool_input:
                            tool_input = tool_input.replace("```json\n", "").replace("\n```", "")
                        
                        input_data = json.loads(tool_input)
                        sheet_name = input_data["sheet_name"]
                        query = input_data["query"]

                        st.write(f"🔍 Analyzing sheet: **{sheet_name}**...")
                        
                        df = excel_db.get_dataframe(sheet_name)
                        if df.empty:
                            return f"The sheet '{sheet_name}' is empty and cannot be analyzed."

                        # --- FIX: Added a more forceful prefix to prevent hallucination ---
                        specialist_prefix = """
                        You are working with a pandas dataframe in Python. The name of the dataframe is `df`.
                        You must use the tools below to answer the question.
                        
                        CRITICAL INSTRUCTION: After you get a result from a tool in the "Observation" step,
                        you MUST use that exact result to form your "Final Answer".
                        Do not make up a new number or value. Your Final Answer must be based directly on the Observation.
                        
                        Example:
                        ...
                        Action: python_repl_ast
                        Action Input: print(df['revenue'].sum())
                        Observation: 3246095
                        Thought: I now know the final answer.
                        Final Answer: The total revenue is 3246095.
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
                        return result["output"]

                    except json.JSONDecodeError:
                        return "Error: The input for the analyze_sheet tool was not valid JSON. Please provide a valid JSON string with 'sheet_name' and 'query' keys."
                    except Exception as e:
                        return f"An error occurred during analysis: {e}"

                def list_sheets_tool(ignored_input=""):
                    """A wrapper for the list_sheets method that ignores any input."""
                    return excel_db.list_sheets()

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
                        name="analyze_sheet",
                        func=analyze_sheet,
                        description="Use this tool to perform detailed analysis, filtering, or calculations on a specific sheet. The input must be a JSON string with 'sheet_name' and 'query' keys."
                    )
                ]
                
                system_message = """
                You are a helpful and friendly "Manager" agent for an Excel file.
                Your primary goal is to delegate tasks to the correct tool based on the user's query.

                **Your Workflow:**
                1.  **Analyze the User's Request:** Carefully read the user's question to understand their intent.
                2.  **Select the Right Tool:**
                    * For general questions about the file's structure, like "What sheets are in this file?" or "How many tables are there?", use the `list_sheets` tool.
                    * To understand the columns of a specific sheet, use `get_sheet_schema`.
                    * For any question that requires looking at the data (e.g., filtering, calculating, finding totals, asking 'how many', 'what is the total', etc.), you **MUST** use the `analyze_sheet` tool.
                3.  **Synthesize the Final Answer:** After using a tool, you must provide a clear, final answer. For example, if the user asks "how many tables are there?", and the `list_sheets` tool returns `['Sales', 'Customers', 'Product Info']`, your final answer must be "There are 3 tables." You must count the items in the list.
                4.  **Format the `analyze_sheet` Input:** When using `analyze_sheet`, your input **must** be a valid JSON object with two keys: "sheet_name" and "query".
                For example, if the user asks "What is the total revenue in the Sales sheet?", you must call the `analyze_sheet` tool and your Action Input would be a JSON object where the `sheet_name` key is "Sales" and the `query` key is "What is the total revenue?".

                **Crucial Rule:** Do **NOT** attempt to answer data-related questions yourself. Your only job is to call the correct tool and then provide a final answer based on the tool's observation.
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
                st.sidebar.success("Analyst Agent Initialized!")
                if not st.session_state.messages:
                    st.session_state.messages.append({"role": "assistant", "content": "Hello! I'm ready to analyze your Excel file. What would you like to know?"})

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
