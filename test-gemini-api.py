import os
import re
import sys
from google import genai
from google.genai import types

def color(text, code):
    if not sys.stdout.isatty() or "NO_COLOR" in os.environ:
        return text
    return f"\033[{code}m{text}\033[0m"

# Check if Windows environment has the key
key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
print("API Key Detected:", color("YES", "92") if key else color("NO", "91"))

# Test Gemini client connection
if key:
    try:
        client = genai.Client()
        response = client.models.generate_content(
            model="gemini-3.7-flash-lite",
            contents="Hello! Confirm connection. Also, how much is 2X2X5?",
            config=types.GenerateContentConfig(
                temperature=0.7,
            )
    )
        print(color("Gemini request succeeded", "92"))
        print("Gemini Response:", response.text)
    except Exception as error:
        status_code = getattr(error, "status_code", None)
        error_text = str(error)
        if status_code is None:
            code_match = re.search(r"['\"]code['\"]\s*:\s*(\d+)|\b(?:HTTP\s+)?(\d{3})\b", error_text, re.IGNORECASE)
            status_code = int(code_match.group(1) or code_match.group(2)) if code_match else "UNKNOWN"
        print(color("Gemini request failed without a traceback.", "91"))
        print(color(f"Error type: {type(error).__name__}", "91"))
        print(color(f"Error code: {status_code}", "91"))
        print(color(f"Error details: {error_text}", "2"))