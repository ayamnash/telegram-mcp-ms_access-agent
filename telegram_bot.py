# telegram_bot.py
import asyncio
import signal
from datetime import datetime

from telegram import Update
from telegram.ext import (
    ApplicationBuilder, CommandHandler,
    MessageHandler, filters, ContextTypes
)

import config_hggingface as config
from mcp_client import MCPManager
from db_context import DB, get_insert_columns
from ai_agent import parse_invoice, missing_fields, build_select_query, format_result

mcp = MCPManager()

# session: user_id → {"conversation": str, "data": dict}
_sessions: dict[int, dict] = {}


# ═══════════════════════════════════════════════
# Intent detection  (simple keywords)
# ═══════════════════════════════════════════════
def _is_invoice(text: str) -> bool:
    t = text.lower()
    return "invoice" in t and any(w in t for w in ["insert", "add", "new", "create"])

def _is_account(text: str) -> bool:
    t = text.lower()
    return any(w in t for w in ["add account", "new customer", "new supplier",
                                 "add customer", "add supplier", "insert account"])

def _is_item(text: str) -> bool:
    t = text.lower()
    return any(w in t for w in ["add item", "new item", "insert item",
                                 "add product", "insert product"])


# ═══════════════════════════════════════════════
# Bot start / stop
# ═══════════════════════════════════════════════
async def run():
    await mcp.connect()

    app = ApplicationBuilder().token(config.BOT_TOKEN).build()
    app.add_handler(CommandHandler("cancel", cmd_cancel))
    app.add_handler(CommandHandler("tools",  cmd_tools))
    app.add_handler(CommandHandler("schema", cmd_schema))
    app.add_handler(CommandHandler("debug",  cmd_debug))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_message))

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


# ═══════════════════════════════════════════════
# Main message handler
# ═══════════════════════════════════════════════
async def on_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.message.from_user.id
    text = update.message.text.strip()

    if text.lower() in ("cancel", "stop"):
        _sessions.pop(uid, None)
        await update.message.reply_text("Cancelled.")
        return

    try:
        # ── Active invoice session ──
        if uid in _sessions:
            await _continue_invoice(update, uid, text)
            return

        # ── New invoice ──
        if _is_invoice(text):
            await _start_invoice(update, uid, text)
            return

        # ── Account insert ──
        if _is_account(text):
            await _insert_account(update, text)
            return

        # ── Item insert ──
        if _is_item(text):
            await _insert_item(update, text)
            return

        # ── Default: SELECT query ──
        await _run_query(update, text)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        await update.message.reply_text(f"Unexpected error: {e}")


# ═══════════════════════════════════════════════
# Invoice workflow
# ═══════════════════════════════════════════════
async def _start_invoice(update: Update, uid: int, text: str):
    """Start a new invoice session."""
    data = parse_invoice(text)
    if data is None:
        await update.message.reply_text("Could not understand the request. Please try again.")
        return

    miss = missing_fields(data)
    if miss:
        _sessions[uid] = {"conversation": text, "data": data}
        await update.message.reply_text(
            "Need a few more details:\n" +
            "\n".join(f"• {m}" for m in miss) +
            "\n\nJust reply with the missing info.\nExample: ayman credit sugar 10*10"
        )
    else:
        await _execute_invoice(update, uid, data)


async def _continue_invoice(update: Update, uid: int, text: str):
    """Continue an existing invoice session with follow-up message."""
    session = _sessions[uid]
    # Append new message to conversation and re-parse everything
    session["conversation"] += f"\n{text}"

    data = parse_invoice(session["conversation"])
    if data is None:
        await update.message.reply_text("Could not parse your reply. Please try again.")
        return

    miss = missing_fields(data)
    if miss:
        session["data"] = data
        await update.message.reply_text(
            "Still need:\n" + "\n".join(f"• {m}" for m in miss)
        )
    else:
        _sessions.pop(uid, None)
        await _execute_invoice(update, uid, data)


async def _execute_invoice(update: Update, uid: int, data: dict):
    """Run the full 5-step invoice insert."""
    await update.message.reply_text("Inserting invoice...")

    # Step 1: resolve AccID
    acc_name = data["acc_name"].strip()
    acc_result = await mcp.call("run_query", {
        "db_name": DB,
        "sql": f"SELECT AccID FROM acctable WHERE AccName = '{acc_name}'"
    })
    acc_id = _first_number(acc_result)
    if acc_id is None:
        await update.message.reply_text(
            f"❌ Account '{acc_name}' not found.\n"
            f"Add them first: add customer {acc_name}"
        )
        return

    # Step 2: resolve ItemIDs + build trans rows
    items      = data.get("items", [])
    tax_pct    = float(data.get("tax_pct") or 0)
    trans_rows = []
    total      = 0.0

    for item in items:
        iname = item["name"].strip()
        qty   = float(item["qty"])
        price = float(item["price"])

        item_result = await mcp.call("run_query", {
            "db_name": DB,
            "sql": f"SELECT ItemID FROM items WHERE ItemName = '{iname}'"
        })
        item_id = _first_number(item_result)
        if item_id is None:
            await update.message.reply_text(
                f"❌ Item '{iname}' not found in items table.\n"
                f"Add it first: add item {iname}"
            )
            return

        line_total = round(qty * price * (1 + tax_pct / 100), 2)
        total += line_total
        trans_rows.append({
            "ItemID":    int(item_id),
            "Qty":       qty,
            "UnitPrice": price,
            "TaxPct":    tax_pct,
            # LineTotal is calculated by Access — do NOT include it
        })

    # Step 3: insert invoice header
    invoice_date = data.get("invoice_date") or datetime.today().strftime("%Y-%m-%d")
    inv_row = {
        "InvoiceDate": invoice_date,
        "AccID":       int(acc_id),
        "InvType":     data.get("inv_type", "sale"),
        "PayType":     data["pay_type"],
        "TotalAmount": round(total, 2),
    }
    if data.get("notes"):
        inv_row["Notes"] = data["notes"]

    ins_result = await mcp.call("insert_data", {
        "db_name": DB, "table": "invoices", "rows": [inv_row]
    })
    if _is_error(ins_result):
        await update.message.reply_text(f"❌ Invoice header insert failed:\n{ins_result}")
        return

    # Step 4: get new InvoiceID
    id_result = await mcp.call("run_query", {
        "db_name": DB,
        "sql": "SELECT TOP 1 InvoiceID FROM invoices ORDER BY InvoiceID DESC"
    })
    invoice_id = _first_number(id_result)
    if invoice_id is None:
        await update.message.reply_text("❌ Could not retrieve new InvoiceID.")
        return

    # Step 5: insert itemstrans
    for row in trans_rows:
        row["InvoiceID"] = int(invoice_id)

    trans_result = await mcp.call("insert_data", {
        "db_name": DB, "table": "itemstrans", "rows": trans_rows
    })
    if _is_error(trans_result):
        await update.message.reply_text(
            f"❌ Invoice header saved (ID={invoice_id}) but items failed:\n{trans_result}"
        )
        return

    # Success
    items_text = "\n".join(
        f"  • {it['name']}  qty={it['qty']}  price={it['price']}"
        for it in items
    )
    await update.message.reply_text(
        f"✅ Invoice inserted!\n\n"
        f"Invoice ID : {invoice_id}\n"
        f"Date       : {invoice_date}\n"
        f"Customer   : {acc_name}\n"
        f"Payment    : {data['pay_type']}\n"
        f"Items:\n{items_text}\n"
        f"Total      : {round(total, 2)}"
    )


# ═══════════════════════════════════════════════
# Account insert
# ═══════════════════════════════════════════════
async def _insert_account(update: Update, text: str):
    tlow = text.lower()
    acc_type = (
        "customer" if "customer" in tlow else
        "supplier" if "supplier" in tlow else
        None
    )
    if not acc_type:
        await update.message.reply_text("Is this a customer or supplier?")
        return

    # Extract name: everything that's not a keyword
    skip = {"add", "insert", "new", "account", "customer", "supplier", "for", "named", "name"}
    import re
    words = [w for w in re.findall(r"[A-Za-z]+", text) if w.lower() not in skip]
    acc_name = " ".join(words) if words else None

    if not acc_name:
        await update.message.reply_text("Please provide the account name.")
        return

    result = await mcp.call("insert_data", {
        "db_name": DB,
        "table":   "acctable",
        "rows":    [{"AccName": acc_name, "AccType": acc_type}]
    })

    if _is_error(result):
        await update.message.reply_text(f"❌ Failed:\n{result}")
    else:
        await update.message.reply_text(f"✅ Account added!\nName: {acc_name}\nType: {acc_type}")


# ═══════════════════════════════════════════════
# Item insert
# ═══════════════════════════════════════════════
async def _insert_item(update: Update, text: str):
    from ai_agent import _chat, _extract_json
    import re

    raw = _chat(
        "Extract item data. Return only JSON.",
        f"""From: "{text}"
Return: {{"ItemCode":"auto like ITM001","ItemName":"name","Unit":"kg/pcs/ltr","SalePrice":number_or_null}}"""
    )
    import json
    try:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        row = json.loads(m.group(0)) if m else json.loads(raw)
    except Exception:
        await update.message.reply_text("Could not parse item. Try: add item sugar unit kg price 2")
        return

    if row.get("SalePrice") is None:
        row.pop("SalePrice", None)

    result = await mcp.call("insert_data", {"db_name": DB, "table": "items", "rows": [row]})

    if _is_error(result):
        await update.message.reply_text(f"❌ Failed:\n{result}")
    else:
        await update.message.reply_text(
            f"✅ Item added!\n"
            f"Name: {row.get('ItemName')}\n"
            f"Code: {row.get('ItemCode')}\n"
            f"Unit: {row.get('Unit')}"
        )


# ═══════════════════════════════════════════════
# SELECT query
# ═══════════════════════════════════════════════
async def _run_query(update: Update, text: str):
    sql = build_select_query(text)
    if not sql:
        await update.message.reply_text("Could not build query. Please rephrase.")
        return

    result = await mcp.call("run_query", {"db_name": DB, "sql": sql})
    nice   = format_result(text, str(result))
    await update.message.reply_text(nice)


# ═══════════════════════════════════════════════
# Commands
# ═══════════════════════════════════════════════
async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    _sessions.pop(update.message.from_user.id, None)
    await update.message.reply_text("Cancelled.")

async def cmd_tools(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    tools = await mcp.list_tools()
    lines = [f"{i}. {t.name}" for i, t in enumerate(tools, 1)]
    await update.message.reply_text("Tools:\n" + "\n".join(lines))

async def cmd_schema(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    table = " ".join(ctx.args) if ctx.args else None
    if not table:
        await update.message.reply_text("Usage: /schema <table>")
        return
    r = await mcp.call("run_query", {"db_name": DB, "sql": f"SELECT TOP 1 * FROM {table}"})
    await update.message.reply_text(str(r))

async def cmd_debug(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = " ".join(ctx.args) if ctx.args else ""
    if not text:
        await update.message.reply_text("Usage: /debug <invoice message>")
        return
    data = parse_invoice(text)
    miss = missing_fields(data) if data else ["parse failed"]
    await update.message.reply_text(f"Parsed:\n{data}\n\nMissing: {miss or 'Nothing ✅'}")


# ═══════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════
def _first_number(query_result) -> int | None:
    """Extract first integer from run_query result string."""
    import re
    text = str(query_result)
    if "no results" in text.lower() or "0 rows" in text.lower():
        return None
    for line in text.splitlines():
        line = line.strip()
        if not line or "---" in line or "Query Results" in line:
            continue
        if line.replace(" ", "").isalpha():
            continue  # column header
        m = re.search(r'\b(\d+)\b', line)
        if m:
            return int(m.group(1))
    return None

def _is_error(result) -> bool:
    text = str(result).lower()
    return any(w in text for w in [
        "error", "exception", "failed", "invalid",
        "unknown", "syntax", "odbc", "hys"
    ])


if __name__ == "__main__":
    asyncio.run(run())