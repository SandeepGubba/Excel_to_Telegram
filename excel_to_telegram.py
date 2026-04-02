import json
import logging
import openpyxl
import os
import prettytable as pt
import sys
from telethon import TelegramClient
import telethon.sync
import asyncio

# pip install openpyxl prettytable telethon

name="excel_to_telegram"
xlsm="excel_to_telegram.xlsm"

# Parameters - will be taken from the Excel params sheet
API_ID=""
API_HASH=""
Channel=""

# Change to the script directory
os.chdir(sys.path[0])

logging.basicConfig(
  filename='excel_to_telegram.log',
  # encoding='utf-8',
  format='%(asctime)s %(levelname)s:%(message)s',
  level=logging.DEBUG
)
logging.debug("Logging activated")

# Log Command line parameters
logging.debug(f"args:{len(sys.argv)}")
for arg in sys.argv:
  logging.debug(f"arg:{arg}")

# Get the data
data=sys.argv[1]
logging.debug(f"data: {data}")
print(f"data: {data}")

def create_table(cells):
    logging.debug(f"create_table - cells: {cells}")
    if not cells:
        return "Empty Data"

    # Use the first row values as headers
    first_row = cells[0]
    headers = []
    for k in first_row:
        val = str(first_row[k]).strip()
        # If the cell is empty, use the column ID (e.g., 'c5') so it doesn't crash
        headers.append(val if val else k) 

    table = pt.PrettyTable(headers)
    
    # Set alignment safely for each header
    for h in headers:
        try:
            table.align[h] = 'l'
        except:
            pass

    # Add the data rows
    for i, r in enumerate(cells):
        if i == 0: continue # Skip the header row (already used for headers)
        row_data = []
        for k in r:
            row_data.append(r[k])
        
        # Only add the row if it matches the header length
        if len(row_data) == len(headers):
            table.add_row(row_data)
            
    return table
# create_table

def get_param(name,wb):
  v=""
  dest = wb.defined_names[name].destinations
  for title,coord in dest:
    range = wb[title][coord]
    v=range.value
  return v
# get_param

def get_params():
  global API_ID
  global API_HASH
  global Channel
  logging.debug(f"get_params from {xlsm}")
  wb = openpyxl.load_workbook(xlsm)
  API_ID=get_param('API_ID',wb)
  API_HASH=get_param('API_HASH',wb)
  Channel=get_param('Channel',wb)
  # logging.debug(f"Params: {API_ID} {API_HASH} {Channel}")
  # print(f"Params: {API_ID} {API_HASH} {Channel}")
  wb.close()
# get_params

async def main():
    # 1. Load Parameters from Excel
    get_params()

    # 2. Setup the Telegram Client
    client = TelegramClient(name, API_ID, API_HASH)
    
    try:
        logging.debug("Starting Telegram Client...")
        await client.start()
          
        # Check if we have data from Excel (sys.argv[1])
        if len(sys.argv) > 1 and sys.argv[1].lower() != 'init':
            # Convert the Excel string into a Python list
            cells = json.loads(sys.argv[1])

            # Build the table using the function from Step 1
            table = create_table(cells)
            
            # Find the Telegram Channel
            entity = await client.get_entity(Channel)
            
            # Send the message
            await client.send_message(entity, f'<pre>{table}</pre>', parse_mode='html')
            logging.debug("Message sent successfully!")
            print("Done! Check your Telegram channel.")
            
            # Short pause to ensure the message is fully sent before disconnecting
            await asyncio.sleep(1)
        else:
            print("Logged in successfully! No data sent (init mode).")
        
    except Exception as e:
        logging.error(f"Error occurred: {e}")
        print(f"Error: {e}")
    finally:
        await client.disconnect()

if __name__ == '__main__':
    # This line is the 'engine' that runs everything above
    asyncio.run(main())
