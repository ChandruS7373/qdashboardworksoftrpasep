# Qualesce AI Project Manager Agent

A Python/Streamlit dashboard with AI agent powered by Claude.

## Run Locally

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.
Enter your Anthropic API key in the sidebar.

## Deploy to Streamlit Cloud (Free)

1. Push this folder to a GitHub repo
2. Go to https://share.streamlit.io
3. Connect your GitHub repo → select app.py
4. Add secret: ANTHROPIC_API_KEY = your_key
5. Click Deploy
