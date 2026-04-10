import os
from dotenv import load_dotenv
from google import genai
from pydantic import BaseModel
from typing import List, Optional

# Load env variables
dotenv_path = os.path.join(os.path.dirname(__file__), "..", ".env")
load_dotenv(dotenv_path)

def get_client() -> genai.Client:
    api_key = os.getenv("gemini_key")
    if not api_key:
        raise ValueError("Missing 'gemini_key' in .env file.")
    return genai.Client(api_key=api_key)
