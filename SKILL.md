# DATABASE SKILL — Invoice System
# Database: K:\MCP_server_ms_access_control-main\invoice.accdb

---

## TOOLS
- insert_data(db_name, table, rows)
- run_query(db_name, sql)

---

## OUTPUT RULE (CRITICAL)
You MUST return ONLY valid JSON.
No text, no explanation, no markdown.
ONLY one JSON object.

---

## RULES
1. NEVER insert AUTO columns:
   - InvoiceID
   - TransID
   - AccID (ONLY auto in acctable)
   - ItemID (ONLY auto in items)
   - LineTotal

2. Numbers must be numbers not strings: 1 not "1"
3. Dates: YYYY-MM-DD format (or #YYYY-MM-DD# for Access SQL queries). Let user provide the valid date; NEVER hallucinate it!
4. NEVER insert LineTotal (calculated by Access)
5. Always use UCASE() for name lookups:
   WHERE UCASE(AccName) = UCASE('ahmad')
6. No @variables in SQL
7. No MSysObjects queries

---

## ACTIONS

Ask user:
{"action": "ask", "message": "..."}

Run query:
{"action": "run_query", "sql": "SELECT ..."}

Insert:
{"action": "insert_data", "table": "items", "rows": [{"ItemName": "sugar"}]}

Done:
{"action": "done", "message": "..."}

Cancel:
{"action": "cancel", "message": "..."}

---

## INTENT RECOGNITION (CRITICAL)
Before doing any workflow, decide what the user is asking. DO NOT mix workflows.

### Intent A: Add Customer Only
If user asks to add or create a customer (e.g. "add new customer faris for acctable table"):
1. Find if customer exists (`run_query`).
2. If not found, insert into `acctable`.
3. Output: `{"action": "done", "message": "Customer added successfully."}`
STOP HERE. DO NOT ask for invoice details. DO NOT mention invoices.

### Intent B: Add Item Only
If user asks to add an item (e.g. "insert new items choco..."):
1. Find if item exists (`run_query`).
2. If not found, insert into `items`.
3. Output: `{"action": "done", "message": "Item added successfully."}`
STOP HERE. DO NOT ask for invoice details. DO NOT mention invoices.

### Intent C: Insert Invoice
Only launch this if the user uses keywords like "insert invoice" or "create invoice".

#### Step 0 — Requirements check for Invoices
You MUST have all of the following explicitly stated in the user prompt. DO NOT proceed until you have explicitly received ALL THREE:
1. **InvoiceDate**: Examine the user's message. Does it contain an explicit, intentional date string (e.g., "2026-05-01", "5/1/2026", "Jan 5")?
   - If NO, you MUST ask for it: `{"action":"ask","message":"Please provide the exact invoice date."}`
   - CRITICAL FIX: DO NOT extract a date from prices, quantities, or total amounts (e.g., if you see "200*10", "10" is a price or quantity, NOT a date!).
   - CRITICAL FIX: DO NOT invent, hallucinate, or assume a date like "1-10-2025" or the current date. 
2. **Customer Name**: If missing → ask: `{"action":"ask","message":"Please provide the customer name."}`
3. **Items, Qty, and Price**: Evaluate the prompt! If the user ALREADY provided the items, quantities, and prices (e.g., "sugar 200*10"), DO NOT ask for them again. ONLY ask if they are truly missing: `{"action":"ask","message":"Please provide items, quantities, and prices."}`

If ANY of these 3 requirements are missing, you MUST ask the user and WAIT. DO NOT execute ANY queries or inserts until you have all 3.

#### Step 1 — Find AccID
`SELECT AccID FROM acctable WHERE UCASE(AccName)=UCASE('ahmad')`
- If not found: ask the user if they want to add it.
- If user says yes: use `insert_data` into the `acctable` table. NEVER invent a table name.

#### Step 2 — Find ItemID
`SELECT ItemID FROM items WHERE UCASE(ItemName)=UCASE('sugar')`
- If not found: ask the user if they want to add it.
- If user says yes: use `insert_data` into the `items` table. NEVER invent a table name.

#### Step 3 — Insert invoice
Use `insert_data` into the table: `invoices`.
(ONLY use the table name `invoices`. Any other name is strictly forbidden.)

#### Step 4 — Get InvoiceID
`SELECT TOP 1 InvoiceID FROM invoices ORDER BY InvoiceID DESC`

#### Step 5 — Insert items
Use `insert_data` into the table: `itemstrans` using the InvoiceID from Step 4.
(ONLY use the table name `itemstrans`. Any other name is strictly forbidden.)
(NEVER include LineTotal)

#### Step 6 — Finish
`{"action": "done", "message": "Invoice #... inserted successfully with items ..."}`

---

## TABLES

### acctable
AccID AUTO  
AccName Text REQUIRED  
AccType Text REQUIRED (customer / supplier / other)  
Phone Text OPTIONAL  

---

### items
ItemID AUTO  
ItemCode Text REQUIRED  
ItemName Text REQUIRED  
Unit Text REQUIRED  
SalePrice Number OPTIONAL  

---

### invoices
InvoiceID AUTO  
InvoiceDate Date REQUIRED  
AccID Number REQUIRED  
InvType Text REQUIRED  
PayType Text REQUIRED  
TotalAmount Number REQUIRED  
Notes Text OPTIONAL  

---

### itemstrans
TransID AUTO  
InvoiceID Number REQUIRED  
ItemID Number REQUIRED  
Qty Number REQUIRED  
UnitPrice Number REQUIRED  
TaxPct Number OPTIONAL  
LineTotal CALCULATED (NEVER INSERT)
