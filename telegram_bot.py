# telegram_bot_v2.py
#
# ARCHITECTURE: Step-by-step loop
#
#   1. User sends message
#   2. Bot sends message + context to AI
#   3. AI returns ONE action: run_query / insert_data / ask / done / cancel
#   4. Bot executes that action
#   5. Bot sends result back to AI: "result was X, next step?"
#   6. Repeat until done or cancel
#
# To switch databases: replace SKILL.md only.

import asyncio
import signal
import json
import re ,time
from pathlib import Path

from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

import config_huggingface as config
from mcp_client import MCPManager
from openai import OpenAI

# ── Load SKILL.md once at startup ──────────────────────────
SKILL = Path("SKILL.md").read_text(encoding="utf-8")
DB    = config.DEFAULT_DB
mcp   = MCPManager()

client = OpenAI(
    base_url=config.HUGGINGFACE_BASE_URL,
    api_key=config.HUGGINGFACE_API_KEY,
)

MODEL = config.HUGGINGFACE_MODEL
MAX_STEPS = 15


# ── Session storage ──────────────────────────────────────
# uid → {"messages": [...], "summary": [...]}
_sessions: dict[int, dict] = {}


# ════════════════════════════════════════════════════════
# SYSTEM PROMPT
# ════════════════════════════════════════════════════════
SYSTEM_PROMPT = f"""You are a database assistant for Microsoft Access.
Database: {DB}

{SKILL}

════════════════════════════════════════
HOW YOU WORK
════════════════════════════════════════

You work step by step. Each turn you receive:
  USER REQUEST: what the user asked
  STEP RESULT:  result of the last action you took

You respond with EXACTLY ONE JSON action.
After each action you will receive the result and decide the next step.

════════════════════════════════════════
AVAILABLE ACTIONS
════════════════════════════════════════

Run a SELECT:
{{"action": "run_query", "sql": "SELECT AccID FROM acctable WHERE UCASE(AccName) = UCASE('watson')"}}

Insert data:
{{"action": "insert_data", "table": "invoices", "rows": [{{"InvoiceDate": "2025-03-03", "AccID": 7, "InvType": "sale", "PayType": "credit", "TotalAmount": 100}}]}}

Ask user for missing info:
{{"action": "ask", "message": "What is the invoice date?"}}

Operation complete:
{{"action": "done", "message": "Invoice #15 inserted. sugar x10 @ 10."}}

Cancel operation:
{{"action": "cancel", "message": "Operation cancelled."}}

════════════════════════════════════════
RULES
════════════════════════════════════════

1. Return ONLY the JSON. No text. No markdown.
2. ONE action per response.
3. Numbers must be numbers: 100 not "100"
4. NEVER insert: InvoiceID, TransID, AccID, ItemID, LineTotal (AUTO or CALCULATED)
5. UCASE for all name lookups: UCASE(AccName) = UCASE('name')
6. If lookup returns no result → ask user "X not found. Add it? (yes/no)"
7. If user says yes → insert, then re-run the lookup to get the new ID
8. If user says no → return cancel action
9. When all steps are done -> return done action with a summary
10. NEVER insert an invoice without first asking the user for items, quantities, and prices if they are missing."""


# ════════════════════════════════════════════════════════
# ONE AI CALL
# ════════════════════════════════════════════════════════
def call_ai(messages):

    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=messages[-12:],  # 🔥 limit history (token optimization)
                temperature=0.1,
                max_tokens=600,
            )

            if not resp.choices:
                raise Exception("Empty response")

            raw = (resp.choices[0].message.content or "").strip()

            print(f"[AI] {raw[:150]}")

            return parse_json(raw)

        except Exception as e:
            print(f"AI error {attempt+1}: {e}")
            time.sleep(1)

    return {"action": "cancel", "message": "AI failed"}

# ===============================
# PARSER (STRONGER)
# ===============================
def parse_json(raw):

    # remove markdown
    m = re.search(r"```(?:json)?\s*(.*?)\s*```", raw, re.DOTALL)
    if m:
        raw = m.group(1)

    # direct
    try:
        return json.loads(raw)
    except:
        pass

    # extract JSON
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except:
            pass

    return {"action": "ask", "message": raw}


# ════════════════════════════════════════════════════════
# EXECUTE ONE ACTION
# ════════════════════════════════════════════════════════
async def execute(action: dict) -> str:
    kind = action.get("action", "")

    if kind == "run_query":
        sql = _fix_ucase(action.get("sql", ""))
        r   = await mcp.call("run_query", {"db_name": DB, "sql": sql})
        return str(r)

    if kind == "insert_data":
        rows = [_sanitize(r) for r in action.get("rows", [])]
        r    = await mcp.call("insert_data", {
            "db_name": DB,
            "table":   action.get("table", ""),
            "rows":    rows,
        })
        return str(r)

    return ""


# ════════════════════════════════════════════════════════
# STEP LOOP — runs until done/cancel/ask
# ════════════════════════════════════════════════════════
async def step_loop(update: Update, uid: int):
    session  = _sessions[uid]
    messages = session["messages"]
    summary  = session["summary"]

    for step in range(MAX_STEPS):
        action = await asyncio.to_thread(call_ai, messages)
        kind   = action.get("action", "reply")
        print(f"  Step {step+1}: {kind}")

        # ── ask user ──────────────────────────────────
        if kind == "ask":
            q = action.get("message", "Please provide more info.")
            await update.message.reply_text(q)
            # Keep session alive — user must reply
            messages.append({"role": "assistant", "content": json.dumps(action)})
            session["waiting"] = True
            return

        # ── done ──────────────────────────────────────
        if kind == "done":
            msg = action.get("message", "Done!")
            if summary:
                msg += "\n" + "\n".join(summary)
            await update.message.reply_text(f"✅ {msg}")
            _sessions.pop(uid, None)
            return

        # ── cancel ────────────────────────────────────
        if kind == "cancel":
            await update.message.reply_text(
                "❌ " + action.get("message", "Cancelled.")
            )
            _sessions.pop(uid, None)
            return

        # ── plain reply ───────────────────────────────
        if kind == "reply":
            await update.message.reply_text(action.get("message", ""))
            _sessions.pop(uid, None)
            return

        # ── execute run_query or insert_data ──────────
        if kind in ("run_query", "insert_data"):
            result = await execute(action)
            print(f"  Result: {result[:150]}")

            if _is_error(result):
                await update.message.reply_text(f"❌ Step {step+1} failed:\n{result}")
                _sessions.pop(uid, None)
                return

            if kind == "insert_data":
                summary.append(f"  • Inserted into {action.get('table','')}")

            # Feed result back to AI
            messages.append({"role": "assistant", "content": json.dumps(action)})
            messages.append({
                "role": "user",
                "content": (
                    f"Step {step+1} result: {result}\n\n"
                    f"What is the next step? "
                    f"When all steps are complete return the 'done' action."
                ),
            })
            continue

        # Unknown
        await update.message.reply_text(f"Unknown action '{kind}'")
        _sessions.pop(uid, None)
        return

    await update.message.reply_text("❌ Too many steps. Please try again.")
    _sessions.pop(uid, None)


# ════════════════════════════════════════════════════════
# TELEGRAM HANDLER
# ════════════════════════════════════════════════════════
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    uid  = update.effective_user.id
    text = update.message.text.strip()
    if not text:
        return

    # Cancel
    if text.lower() in ("cancel", "stop", "/cancel", "/start"):
        _sessions.pop(uid, None)
        await update.message.reply_text("Session cleared. Start a new request.")
        return

    try:
        # Direct show commands — no AI needed
        if await _handle_show(update, text):
            return

        # Continue existing session (user answered a question)
        if uid in _sessions and _sessions[uid].get("waiting"):
            session = _sessions[uid]
            session["waiting"] = False
            session["messages"].append({
                "role": "user",
                "content": (
                    f"User answered: {text}\n\n"
                    f"Continue. What is the next step?"
                ),
            })
            await step_loop(update, uid)
            return

        # New request — start fresh session
        _sessions[uid] = {
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": f"USER REQUEST: {text}"},
            ],
            "summary": [],
            "waiting": False,
        }
        await step_loop(update, uid)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        _sessions.pop(uid, None)
        await update.message.reply_text(f"❌ Error: {e}")


# ════════════════════════════════════════════════════════
# DIRECT SHOW COMMANDS
# ════════════════════════════════════════════════════════
async def _handle_show(update: Update, text: str) -> bool:
    t = text.lower().strip()

    QUERIES = [
        (
            ["list item","list items","show items",
             "get items","get me items","items list","item list",
             "get me list of item","get me list of items"],
            "SELECT ItemID, ItemName, Unit, SalePrice FROM items"
        ),
        (
            ["list accounts", "list account",
             "show accounts","get accounts","get customers","show customers",
             "get me customers","all customers","get all customers name",
             "customers name","get me all customers name"],
            "SELECT AccID, AccName, AccType FROM acctable"
        ),
        (
            ["list invoices","show invoices","list invoice", "show invoice",
             "get invoices","last invoices","get me invoices"],
            "SELECT TOP 10 InvoiceID, InvoiceDate, AccID, PayType, TotalAmount "
            "FROM invoices ORDER BY InvoiceID DESC"
        ),
    ]

    for keywords, sql in QUERIES:
        if t in keywords:
            try:
                r = await mcp.call("run_query", {"db_name": DB, "sql": sql})
                await update.message.reply_text(str(r))
            except Exception as e:
                import traceback
                traceback.print_exc()
                await update.message.reply_text(f"Query error: {e}")
            return True

    return False


# ════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════
def _is_error(result: str) -> bool:
    low = result.lower()
    return any(w in low for w in [
        "error","exception","failed","invalid",
        "odbc","hys","syntax","too few","mismatch"
    ])


def _fix_ucase(sql: str) -> str:
    text_cols = ["AccName","ItemName","FullName","MedicineName","ProductName"]
    for col in text_cols:
        pattern = r"(?i)\b" + re.escape(col) + r"\s*=\s*'([^']+)'"
        def make_ucase(m, c=col):
            return f"UCASE({c}) = UCASE('{m.group(1)}')"
        if "ucase(" + col.lower() + ")" not in sql.lower():
            sql = re.sub(pattern, make_ucase, sql)
    return sql


def _sanitize(row: dict) -> dict:
    clean = {}
    for k, v in row.items():
        if v is None or v == "":
            clean[k] = v
        elif isinstance(v, (int, float, bool)):
            clean[k] = v
        elif isinstance(v, str):
            s = v.strip()
            try:    clean[k] = int(s)
            except ValueError:
                try:    clean[k] = float(s)
                except ValueError: clean[k] = v
        else:
            clean[k] = v
    return clean


# ════════════════════════════════════════════════════════
# STARTUP
# ════════════════════════════════════════════════════════
async def run():
    await mcp.connect()
    print(f"Model    : {MODEL}")
    print(f"Database : {DB}")
    print(f"SKILL.md : {len(SKILL)} chars")

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