# agent.py

import os
import pandas as pd
from typing import Dict, Union
from openai import OpenAIError

from langchain_openai import ChatOpenAI
from langchain_experimental.agents import create_pandas_dataframe_agent

from excel_utils import read_excel_file


def load_openai_api():
    """Load OpenAI API key from environment (or .env)."""
    from dotenv import load_dotenv
    load_dotenv()
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY not set in environment.")
    return api_key


def create_excel_agent(file_path: str, model_name: str = "gpt-4") -> Union[None, object]:
    try:
        load_openai_api()

        sheets: Dict[str, pd.DataFrame] = read_excel_file(file_path)

        print("\n📄 Excel Sheets Found:")
        for name, df in sheets.items():
            print(f"- {name}: {df.shape[0]} rows, columns: {list(df.columns)}")

        all_dataframes = list(sheets.values())
        llm = ChatOpenAI(model=model_name, temperature=0)

        agent = create_pandas_dataframe_agent(
            llm=llm,
            df=all_dataframes,
            verbose=False,
            agent_type="openai-tools",
            return_intermediate_steps=False,
            allow_dangerous_code=True
        )
        return agent

    except OpenAIError as e:
        print(f"OpenAI error: {e}")
        return None
    except Exception as e:
        print(f"Error creating agent: {e}")
        return None


def run_interactive_chat(agent):
    print("\n🤖 Ask me anything about your Excel data!")
    print("Type 'exit' or 'quit' to stop.\n")

    while True:
        try:
            user_input = input("❓ Your question: ").strip()
            if user_input.lower() in {"exit", "quit"}:
                print("👋 Goodbye!")
                break

            response = agent.invoke({"input": user_input})
            print("\n📊 Answer:\n", response.get("output", "No output returned."))
        except KeyboardInterrupt:
            print("\n👋 Exiting.")
            break
        except Exception as e:
            print(f"⚠️ Error: {e}\n")


if __name__ == "__main__":
    file_path = "C:\\Users\\aravi\\Downloads\\Assignments\\W6D1\\excel-agent\\sales_data.xlsx"

    agent = create_excel_agent(file_path)

    if agent is not None:
        run_interactive_chat(agent)
    else:
        print("❌ Failed to create agent. Check for errors above.")
