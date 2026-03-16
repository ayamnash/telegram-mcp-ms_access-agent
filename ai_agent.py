# ai_agent.py
import json
import re
from openai import OpenAI
import config_hggingface as config
from db_context import schema_summary, DB

client = OpenAI(
    base_url=config.HUGGINGFACE_BASE_URL,
    api_key=config.HUGGINGFACE_API_KEY,
)

# ── low-level chat call ──────────────────────────────────────
def _chat(system: str, user: str, temperature=0.0) -> str:
    resp = client.chat.completions.create(
        model=config.HUGGINGFACE_MODEL,
        messages=[
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        temperature=temperature,
    )
    return (resp.choices[0].message.content or "").strip()

def _extract_json(text: str):
    text = text.strip()
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
    return None


# ── Parse invoice from natural language ─────────────────────
def parse_invoice(conversation: str) -> dict | None:
    """
    Given the full conversation text (all messages so far),
    extract a structured invoice dict.

    Returns:
    {
      "invoice_date": "YYYY-MM-DD",   # or null
      "acc_name":     "ayman",        # customer name string
      "pay_type":     "cash|credit",  # or null
      "inv_type":     "sale",
      "tax_pct":      0,
      "items": [
        {"name": "sugar", "qty": 10, "price": 10}
      ]
    }
    Null fields = still missing.
    """
    system = "You extract invoice data from conversation text. Return only valid JSON."
    user = f"""Extract invoice data from this conversation and return ONE JSON object.

CONVERSATION:
{conversation}

Return exactly this structure (use null for missing fields):
{{
  "invoice_date": "YYYY-MM-DD or null",
  "acc_name":     "customer name or null",
  "pay_type":     "cash or credit or null",
  "inv_type":     "sale",
  "tax_pct":      0,
  "items": [
    {{"name": "item name", "qty": number, "price": number}}
  ]
}}

Rules:
- Dates like "1-2-2026" or "1/2/2026" → "2026-02-01"
- qty*price means quantity * unit_price  (e.g. 10*10 → qty=10, price=10)
- items list must only contain items that have BOTH qty AND price
- Return ONLY the JSON, no explanation"""

    raw = _chat(system, user)
    print(f"DEBUG parse_invoice raw: {raw}")
    return _extract_json(raw)


def missing_fields(data: dict) -> list[str]:
    """Returns human-readable list of what's still needed."""
    missing = []
    if not data.get("acc_name"):
        missing.append("Customer name (e.g. 'for ayman')")
    if not data.get("pay_type"):
        missing.append("Payment type: cash or credit")
    if not data.get("items"):
        missing.append("Items with quantity and price (e.g. 'sugar 10*10')")
    return missing


# ── Build SELECT query ───────────────────────────────────────
def build_select_query(user_text: str) -> str | None:
    """Ask AI to write a SELECT query. Returns SQL string or None."""
    system = "You write Microsoft Access SQL SELECT queries. Return only the SQL string, no JSON, no explanation."
    user = f"""Database schema:
{schema_summary()}

Write a SELECT query for: "{user_text}"

Rules:
- Use exact column/table names from schema
- Access date format: #YYYY-MM-DD#
- Single quotes for strings: 'value'
- Return ONLY the SQL, nothing else"""

    sql = _chat(system, user)
    # Strip any accidental markdown
    sql = re.sub(r"```[a-z]*\n?", "", sql).strip("`").strip()
    print(f"DEBUG build_select_query: {sql}")
    return sql if sql else None


# ── Format result for Telegram ───────────────────────────────
def format_result(user_text: str, raw_result: str) -> str:
    """Format a real DB result into a clean Telegram message."""

    # Never beautify errors
    if any(w in raw_result.lower() for w in [
        "error", "exception", "failed", "odbc", "hys", "syntax"
    ]):
        return f"❌ Operation failed:\n{raw_result}"

    system = (
        "You format database results into clean Telegram messages. "
        "Format ONLY what is in the result. Never invent data."
    )
    user = f"""User request: {user_text}
Database result: {raw_result}

Format into a clean readable Telegram message. Use * for bold. No code blocks."""

    return _chat(system, user, temperature=0.1)