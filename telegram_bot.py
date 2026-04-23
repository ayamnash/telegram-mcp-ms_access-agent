import asyncio
import json
import re
import signal
import time
from pathlib import Path

from openai import OpenAI
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, MessageHandler, filters

import config_huggingface as config
from mcp_client import MCPManager


SKILL = Path("SKILL.md").read_text(encoding="utf-8")
mcp = MCPManager()

client = OpenAI(
    base_url=config.HUGGINGFACE_BASE_URL,
    api_key=config.HUGGINGFACE_API_KEY,
)

MODEL = config.HUGGINGFACE_MODEL
MAX_STEPS = 10
DOMAIN = "accounting"


_sessions: dict[int, dict] = {}


SYSTEM_PROMPT = f"""You are an accounting AI agent.
Your job is to understand the user's request, choose exactly one safe business operation, and build the payload.

{SKILL}

You work in a loop:
1. Read the user request.
2. If data is missing, ask the user.
3. When ready, call exactly one business operation.
4. After the tool result returns, either ask a follow-up question, finish, or cancel.

Return EXACTLY ONE JSON object each time.

Available actions:
{{"action": "execute_operation", "domain": "accounting", "operation": "accounting.create_invoice", "payload": {{"invoice_date": "2026-04-20", "customer_name": "Ahmad", "inv_type": "sale", "pay_type": "cash", "items": [{{"item_name": "Sugar", "quantity": 2, "unit_price": 5}}]}}}}
{{"action": "ask", "message": "What is the invoice date?"}}
{{"action": "done", "message": "I found the latest invoices."}}
{{"action": "cancel", "message": "Operation cancelled."}}

Allowed operations:
- accounting.list_items
- accounting.list_customers
- accounting.list_recent_invoices
- accounting.find_item
- accounting.find_customer
- accounting.find_invoices_by_customer
- accounting.get_invoice_details
- accounting.list_invoices_by_date_range
- accounting.total_sales_per_customer
- accounting.create_customer
- accounting.create_item
- accounting.create_invoice

Strict rules:
1. Return JSON only. No markdown. No extra text.
2. Never write SQL.
3. Never mention table names or column names to the user.
4. For any database action, use only execute_operation.
5. Domain must always be "accounting".
6. Never split invoice creation into header/detail SQL steps. Use accounting.create_invoice once when all required values are ready.
7. If a customer or item does not exist, ask the user first. Only use accounting.create_customer or accounting.create_item after explicit confirmation.
8. If the tool result says status is "needs_input", ask the user for the missing data instead of guessing.
9. After a successful read operation, return done with the answer and the important records.
10. After a successful write operation, return done with a short success summary.
"""


def parse_json(raw: str) -> dict:
    code_block = re.search(r"```(?:json)?\s*(.*?)\s*```", raw, re.DOTALL)
    if code_block:
        raw = code_block.group(1)

    try:
        return json.loads(raw)
    except Exception:
        pass

    object_match = re.search(r"\{.*\}", raw, re.DOTALL)
    if object_match:
        try:
            return json.loads(object_match.group(0))
        except Exception:
            pass

    return {"action": "ask", "message": raw.strip() or "Please clarify your request."}


def parse_tool_result(raw) -> dict:
    if isinstance(raw, dict):
        return raw

    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except Exception:
            return {
                "status": "error",
                "message": f"Invalid tool response: {raw}",
            }

    return {
        "status": "error",
        "message": f"Unsupported tool response type: {type(raw).__name__}",
    }


def call_ai(messages: list[dict]) -> dict:
    for attempt in range(3):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=messages[-14:],
                temperature=0.1,
                max_tokens=900,
            )

            if not response.choices:
                raise RuntimeError("Empty response from model")

            raw = (response.choices[0].message.content or "").strip()
            print(f"[AI] {raw[:200]}")
            return parse_json(raw)
        except Exception as exc:
            print(f"AI error {attempt + 1}: {exc}")
            time.sleep(1)

    return {"action": "cancel", "message": "AI failed to complete the request."}


async def execute_action(action: dict) -> dict:
    result = await mcp.call(
        "execute_operation",
        {
            "operation": action.get("operation", ""),
            "payload": action.get("payload", {}),
            "domain": action.get("domain", DOMAIN),
        },
    )
    return parse_tool_result(result)


async def step_loop(update: Update, uid: int) -> None:
    session = _sessions[uid]
    messages = session["messages"]

    for step in range(MAX_STEPS):
        action = await asyncio.to_thread(call_ai, messages)
        action_type = action.get("action", "")
        print(f"Step {step + 1}: {action_type}")

        if action_type == "ask":
            question = action.get("message", "Please provide the missing information.")
            await update.message.reply_text(question)
            messages.append({"role": "assistant", "content": json.dumps(action, ensure_ascii=False)})
            session["waiting"] = True
            return

        if action_type == "done":
            await update.message.reply_text(action.get("message", "Done."))
            _sessions.pop(uid, None)
            return

        if action_type == "cancel":
            await update.message.reply_text(action.get("message", "Operation cancelled."))
            _sessions.pop(uid, None)
            return

        if action_type == "execute_operation":
            try:
                tool_result = await execute_action(action)
            except Exception as exc:
                await update.message.reply_text(f"Tool execution failed: {exc}")
                _sessions.pop(uid, None)
                return

            messages.append({"role": "assistant", "content": json.dumps(action, ensure_ascii=False)})
            messages.append(
                {
                    "role": "user",
                    "content": (
                        "TOOL RESULT:\n"
                        + json.dumps(tool_result, ensure_ascii=False)
                        + "\n\nDecide the next action. If the result is enough for the user, return done."
                    ),
                }
            )
            continue

        await update.message.reply_text(f"Unknown action '{action_type}'.")
        _sessions.pop(uid, None)
        return

    await update.message.reply_text("Too many steps. Please try again.")
    _sessions.pop(uid, None)


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.text:
        return

    uid = update.effective_user.id
    text = update.message.text.strip()
    if not text:
        return

    if text.lower() in ("cancel", "stop", "/cancel", "/start"):
        _sessions.pop(uid, None)
        await update.message.reply_text("Session cleared. Start a new request.")
        return

    try:
        if uid in _sessions and _sessions[uid].get("waiting"):
            session = _sessions[uid]
            session["waiting"] = False
            session["messages"].append(
                {
                    "role": "user",
                    "content": f"USER ANSWER: {text}",
                }
            )
            await step_loop(update, uid)
            return

        _sessions[uid] = {
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"USER REQUEST: {text}"},
            ],
            "waiting": False,
        }
        await step_loop(update, uid)
    except Exception as exc:
        import traceback

        print(traceback.format_exc())
        _sessions.pop(uid, None)
        await update.message.reply_text(f"Error: {exc}")


async def run() -> None:
    await mcp.connect()
    print(f"Model: {MODEL}")
    print(f"SKILL.md size: {len(SKILL)} chars")

    app = ApplicationBuilder().token(config.BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT, handle))
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    print("BOT RUNNING")

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, stop.set)
        except NotImplementedError:
            pass

    try:
        await stop.wait()
    except KeyboardInterrupt:
        pass
    finally:
        await app.updater.stop()
        await app.stop()
        await app.shutdown()
        await mcp.close()


if __name__ == "__main__":
    asyncio.run(run())
