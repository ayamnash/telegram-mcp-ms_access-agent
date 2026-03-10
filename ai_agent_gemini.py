import json
import re

import google.generativeai as genai

import config_gemini as config

genai.configure(api_key=config.GEMINI_API_KEY)
model = genai.GenerativeModel(config.GEMINI_MODEL)


def extract_json(text):
    """Extract JSON from model response, handling markdown code fences."""
    text = text.strip()

    match = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL)
    if match:
        text = match.group(1).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass

    return {"reply": text}


def chat_completion(messages, temperature=0.2):
    """Send messages to Gemini and return the text response."""
    # Convert OpenAI-style messages to Gemini format
    system_prompt = ""
    gemini_history = []

    for msg in messages:
        role = msg["role"]
        content = msg["content"]

        if role == "system":
            system_prompt = content
        elif role == "user":
            gemini_history.append({"role": "user", "parts": [content]})
        elif role == "assistant":
            gemini_history.append({"role": "model", "parts": [content]})

    # Re-create model with system instruction if present
    active_model = (
        genai.GenerativeModel(
            config.GEMINI_MODEL,
            system_instruction=system_prompt,
        )
        if system_prompt
        else model
    )

    # Last user message is sent as the prompt; history is everything before it
    chat = active_model.start_chat(history=gemini_history[:-1])
    last_user_msg = gemini_history[-1]["parts"][0] if gemini_history else ""

    response = chat.send_message(
        last_user_msg,
        generation_config=genai.types.GenerationConfig(temperature=temperature),
    )
    return response.text.strip()


async def decide_tool(user_text, tools, context_info=None):
    tool_desc = [
        {
            "name": t.name,
            "description": t.description,
            "inputSchema": t.inputSchema,
        }
        for t in tools
    ]

    context_section = ""
    if context_info:
        context_section = f"""
IMPORTANT CONTEXT (use this information):
{context_info}
"""

    prompt = f"""
You control Microsoft Access via MCP tools.
The default database is: {config.DEFAULT_DB}
Always use this database unless the user specifies a different one.

TOOLS:
{json.dumps(tool_desc, indent=2, ensure_ascii=False)}
{context_section}
USER REQUEST:
{user_text}

CRITICAL RULES FOR INSERT OPERATIONS:
- For INSERT operations, you MUST use the EXACT field names provided in the context
- Never guess field names - only use what's given to you
- For insert_data tool, the 'rows' parameter must be a list of dictionaries with exact field names
- Generate realistic sample data with proper data types
- For TEXT fields: use realistic strings
- For DATETIME fields: use format "YYYY-MM-DD" or "YYYY-MM-DD HH:MM:SS"
- For CURRENCY/NUMBER fields: use numeric values (not strings)
- Do NOT include AUTOINCREMENT fields (like InvoiceID) in the insert data
- Example for invoice_tabl:
  {{"InvoiceNumber": "INV-001", "CustomerName": "ABC Corp", "InvoiceDate": "2026-03-10", "DueDate": "2026-04-10", "TotalAmount": 1500.50, "Status": "Pending", "Notes": "Sample invoice"}}

Return ONLY ONE raw JSON object, no markdown, no explanation. Do NOT return a list or array:

{{
 "tool_name": "...",
 "arguments": {{ }}
}}
"""

    txt = chat_completion(
        [
            {
                "role": "system",
                "content": "You select the best MCP tool and must answer with exactly one JSON object. Follow the field names exactly as provided. Generate realistic sample data.",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.3,
    )

    return extract_json(txt)


async def format_response(user_text, raw_result):
    """Use Gemini to format the raw MCP result for Telegram."""

    prompt = f"""
You are a helpful Telegram bot assistant. Format the following database result
into a clean, readable Telegram message.

USER ASKED: {user_text}

RAW RESULT:
{raw_result}

Rules:
- Use emojis to make it visually appealing
- Use clean formatting (bold with *, lists with bullet points)
- Keep it concise and organized
- Do NOT use markdown code blocks
- For table lists, use numbered lists with emojis
- For query results, format as a clean table or list
- Add a short friendly header
- If there's an error, explain it simply
- For INSERT operations that return "Inserted X rows", celebrate the success! ✅
- If result contains "successfully" or "Inserted", make it positive and encouraging
"""

    return chat_completion(
        [
            {
                "role": "system",
                "content": "You format tool outputs into concise Telegram-friendly replies. Make successful operations feel rewarding!",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.4,
    )