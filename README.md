## Demo Video

[![Watch the video](https://img.youtube.com/vi/EWxoyoI3oHY/0.jpg)](https://youtu.be/EWxoyoI3oHY)
## v 0.02
Giving AI information about your database and how to handle it.
AI = brain, Python bot = coordinator, MCP server = hands that touch the database
![Architecture](./images/db_context_v.02.svg)
# Telegram MCP MS Access Agent

Control a **Microsoft Access database** directly from **Telegram** using an **AI Agent** powered by the **Model Context Protocol (MCP)**.

This project demonstrates how natural language messages sent to a Telegram bot can trigger database operations such as inserting records and querying data from a local Microsoft Access database.

The system connects Telegram, an AI model, and a local database through MCP to enable remote database interaction.

---

# Project Repository

GitHub:
https://github.com/ayamnash/telegram-mcp-ms_access-agent

---

# Overview

This project shows how an AI agent can interact with a local database through MCP while using Telegram as a remote user interface.

Instead of manually accessing the database, users simply send commands to a Telegram bot.  
The AI agent interprets the request and executes the required operation on the database.

Currently supported operations:

- Insert records into Microsoft Access
- Query data from the database

---

# Architecture

The system follows a simple architecture:

```

Telegram User
↓
Telegram Bot
↓
AI Agent
↓
MCP Client
↓
MCP Server
↓
Microsoft Access Database

```

MCP acts as the communication layer that allows AI agents to interact with external tools and services in a standardized way. :contentReference[oaicite:1]{index=1}

---

# Example Usage

Insert a record:

```

Add new employee
Name: John
Salary: 500

```

Query data:

```

Show all employees

```

The bot processes the request and sends the database results back to Telegram.

---

# Requirements

Ensure the architecture is consistent between Python and Microsoft Office.

Use either:

### 32-bit Environment

```

Python 32-bit
Microsoft Office 32-bit
Access Database Engine 32-bit

```

or

### 64-bit Environment

```

Python 64-bit
Microsoft Office 64-bit
Access Database Engine 64-bit

```

Mixing architectures may cause database driver errors.

---

# Installation

Clone the repository

```

git clone https://github.com/ayamnash/telegram-mcp-ms_access-agent.git
cd telegram-mcp-ms_access-agent

```

Install dependencies

```

pip install -r requirements.txt

```

---

# Configuration

Create or edit `config.py` and add your credentials:

```

BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN"
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
HUGGINGFACE_API_KEY= "YOUR_HUGGINGFACE_KEY"

```

---

# Run the Bot

Start the application

```

python main.py

```

If everything is configured correctly you should see:

```

BOT RUNNING

```

Now you can control your Microsoft Access database directly from Telegram.

---

# Dependencies

The project uses the following Python libraries:

- fastmcp
- python-telegram-bot
- pyodbc
- pywin32
- google-generativeai
- openai

---

# Use Cases

- Remote database management
- AI-assisted database queries
- Telegram-based automation
- AI agent experimentation with MCP

---

# License

This project is open-source and provided for educational and experimental purposes.
```

---

💡 إذا أردت، أستطيع أيضاً أن .
