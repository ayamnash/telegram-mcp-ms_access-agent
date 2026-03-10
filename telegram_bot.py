import asyncio
import signal

from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

#import config_gemini as config
import config_hggingface as config
from mcp_client import MCPManager
#from ai_agent_gemini import decide_tool , format_response
from ai_agent_huggingface import decide_tool , format_response


mcp = MCPManager()


async def start():

    await mcp.connect()

    app = ApplicationBuilder().token(config.BOT_TOKEN).build()


    app.add_handler(MessageHandler(filters.TEXT, handle))

    # Use the lower-level async API instead of run_polling()
    # run_polling() tries to manage its own event loop, which conflicts
    # with the already-running loop from asyncio.run()
    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    print("BOT RUNNING")

    # Keep running until interrupted
    stop_event = asyncio.Event()

    # Handle Ctrl+C gracefully
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            loop.add_signal_handler(sig, stop_event.set)
        except NotImplementedError:
            # Windows doesn't support add_signal_handler for SIGTERM
            pass

    try:
        await stop_event.wait()
    except KeyboardInterrupt:
        pass
    finally:
        print("Shutting down...")
        await app.updater.stop()
        await app.stop()
        await app.shutdown()
        await mcp.close()


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user_text = update.message.text

    try:
        tools = await mcp.list_tools()

        # Check if this is an INSERT operation - if so, get schema first
        context_info = None
        is_insert_operation = False
        table_name = None
        
        if "insert" in user_text.lower():
            is_insert_operation = True
            # Try to extract table name
            import re
            table_match = re.search(r'(?:into|for)\s+(\w+)', user_text, re.IGNORECASE)
            if table_match:
                table_name = table_match.group(1)
                try:
                    # Get table schema by running a query
                    schema_result = await mcp.call(
                        "mcp_msaccess_run_query",
                        {
                            "db_name": config.DEFAULT_DB,
                            "sql": f"SELECT TOP 1 * FROM {table_name}"
                        }
                    )
                    
                    # Extract column names from the result
                    if schema_result and len(schema_result) > 0:
                        columns = list(schema_result[0].keys())
                        # Filter out AUTOINCREMENT fields (usually end with ID)
                        columns_str = ', '.join(columns)
                        context_info = f"Table '{table_name}' has these exact columns: {columns_str}. IMPORTANT: Do NOT include InvoiceID (it's AUTOINCREMENT). Only use: {', '.join([c for c in columns if 'ID' not in c or c == 'CustomerID'])}."
                except:
                    pass  # If schema fetch fails, continue without it

        # Make decision with or without schema context
        decision = await decide_tool(user_text, tools, context_info)

        # If it returned a list, just take the first tool call
        if isinstance(decision, list):
            if len(decision) > 0:
                decision = decision[0]
            else:
                await update.message.reply_text("I didn't understand what to do.")
                return

        # If the AI returned a clarifying question instead of a tool call
        if "reply" in decision:
            await update.message.reply_text(decision["reply"])
            return

        if "tool_name" not in decision:
            await update.message.reply_text("I couldn't determine the correct tool to use.")
            return

        tool = decision["tool_name"]
        args = decision.get("arguments", {})

        # Call the tool
        result = await mcp.call(tool, args)
        
        # Log the raw result for debugging
        print(f"DEBUG - Tool: {tool}")
        print(f"DEBUG - Args: {args}")
        print(f"DEBUG - Raw result: {result}")
        print(f"DEBUG - Result type: {type(result)}")

        # For insert operations, automatically verify by querying the table
        if is_insert_operation and table_name and "insert" in tool.lower():
            try:
                verify_result = await mcp.call(
                    "mcp_msaccess_run_query",
                    {
                        "db_name": config.DEFAULT_DB,
                        "sql": f"SELECT TOP 5 * FROM {table_name} ORDER BY InvoiceID DESC"
                    }
                )
                result = f"✅ Successfully inserted records!\n\nHere are the latest records:\n{verify_result}"
            except:
                pass

        # Format the ugly MCP result into a nice, readable Telegram message
        nice_text = await format_response(user_text, str(result))

        await update.message.reply_text(nice_text)

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"ERROR: {error_details}")
        await update.message.reply_text(f"Error: {e}")