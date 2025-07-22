# Excel Agent

This project provides an agent for automating tasks with Excel files using OpenAI's GPT-4 model.

## Setup

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Set up your OpenAI API key:
   - Get your API key from [OpenAI Platform](https://platform.openai.com/api-keys)
   - Create a `.env` file in the project root with:
     ```
     OPENAI_API_KEY=your_openai_api_key_here
     ```

3. Run the Streamlit app:
   ```
   streamlit run app.py
   ```

## Usage

1. Upload an Excel file using the file uploader
2. Ask questions about your data in natural language
3. The agent will interpret your question and execute the appropriate commands

## Project Structure
- `app.py`: Streamlit web application entry point
- `agent/`: Core agent implementation
  - `agent.py`: Main agent logic
  - `chains.py`: LangChain processing chains
  - `tools.py`: Excel processing tools
  - `column_mapper.py`: Column mapping utilities
  - `prompts.py`: Prompt templates
- `requirements.txt`: Python dependencies
- `.gitignore`: Files and folders to ignore in git
- `README.md`: Project documentation
