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
3. Dates: YYYY-MM-DD
4. NEVER insert LineTotal (calculated by Access)
5. Always use UCASE() for name lookups:
   WHERE UCASE(AccName) = UCASE('ahmad')

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

---

## INSERT INVOICE WORKFLOW

### Step 0 — Requirements check
You MUST have:
- InvoiceDate
- Customer Name
- At least one item with Qty and Price

If missing → ask:
{"action":"ask","message":"Please provide items, quantities, and prices."}

---

### Step 1 — Find AccID
SELECT AccID FROM acctable WHERE UCASE(AccName)=UCASE('ahmad')

If not found:
→ ask to add

---

### Step 2 — Find ItemID
Same logic as account

---

### Step 3 — Insert invoice
insert_data invoices

---

### Step 4 — Get InvoiceID
SELECT TOP 1 InvoiceID FROM invoices ORDER BY InvoiceID DESC

---

### Step 5 — Insert items
insert_data itemstrans
(NEVER include LineTotal)

---

## ACCESS RULES
- Date format: #2025-03-04#
- No @variables
- No MSysObjects