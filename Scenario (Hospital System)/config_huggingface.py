import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BOT_TOKEN = ""
HUGGINGFACE_API_KEY = ""
HUGGINGFACE_BASE_URL = "https://router.huggingface.co/v1"
HUGGINGFACE_MODEL = "Qwen/Qwen2.5-7B-Instruct"

#GEMINI_API_KEY = "AIzaSyCqU9wJCHpZb-AlHa0FaAmk7Vgv0aqq7q4"
#GEMINI_MODEL   = "gemini-2.5-flash"
# python location   
#Python 32-bit Microsoft Office 32-bit Access Database Engine 32-bit 
#or  64-bit Microsoft Office 64-bit Access Database Engine 64-bit 
MCP_COMMAND = r"C:\Users\it\AppData\Local\Programs\Python\Python311-32\python.exe"
MCP_SCRIPT = os.path.join(BASE_DIR, "server.py")

DEFAULT_DB = os.path.join(BASE_DIR, "Hospital.accdb")
