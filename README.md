# OritDva — Personalized Email Style Responder

An LLM-powered system that **learns your writing style** from sample texts and **drafts email replies** in your voice using Gemini and Outlook.

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Configure
copy .env.example .env
# Edit .env and paste your Gemini API key

# 3. Collect writing samples from Outlook (reads your Sent Items)
python main.py collect

# 4. Extract your style
python main.py extract

# 5. Test connectivity
python main.py test

# 6. Process emails
python main.py respond
```

## Commands

| Command | Description |
|---------|-------------|
| `python main.py collect` | Pull sent emails from Outlook as writing samples |
| `python main.py extract` | Analyze writing samples and create a style profile |
| `python main.py check` | List unread emails (preview only) |
| `python main.py respond` | Generate draft replies for unread emails |
| `python main.py folders` | List available Outlook folders |
| `python main.py test` | Test Outlook & Gemini connectivity |

## How It Works

1. **Style Extraction** — Gemini analyzes your writing samples to build a "Stylistic DNA" profile (tone, vocabulary, greeting patterns, temper, etc.)
2. **Email Reading** — Connects to your local Outlook via COM to fetch unread emails
3. **Reply Generation** — Gemini generates replies that mimic your exact writing style
4. **Draft Creation** — Replies are saved as **drafts** in Outlook for your review

## Requirements

- Windows with Outlook installed and running
- Python 3.10+
- Gemini API key ([get one here](https://aistudio.google.com/apikey))
