"""
Configuration loader for OritDva email responder.
Reads settings from .env file and provides defaults.
"""
import os
from dotenv import load_dotenv

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
OUTLOOK_FOLDER = os.getenv("OUTLOOK_FOLDER", "Inbox")
STYLE_SAMPLES_DIR = os.getenv("STYLE_SAMPLES_DIR", "./samples")
STYLE_PROFILE_PATH = os.getenv("STYLE_PROFILE_PATH", "./style_profile.json")

# Gemini model to use
GEMINI_MODEL = "gemini-2.0-flash"
