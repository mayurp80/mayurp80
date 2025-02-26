import asyncio
import websockets
import json
import xlwings as xw
import pandas as pd
import time

# Open Excel Workbook (Keeps it Open)
wb = xw.Book("data.xlsx")

async def save_excel(sheet):
    """Save Excel only when A2 == 1"""
    while True:
        try:
            if sheet.range("A2").value == 1:
                wb.save()
                print("💾 Auto-saved data.xlsx")
        except Exception as e:
            print("⚠️ Error saving Excel:", str(e))
        await asyncio.sleep(6)  # Check every 5 seconds

async def send_data(websocket, sheet):
    """Send data when A2 == 1"""
    try:
        while True:
            if sheet.range("A2").value == 1:
                attempt = 0
                # Set B2 to 1 before starting the data reading process
                sheet.range("B2").value = 1
                print("💬 Reading data...")

                while attempt < 5:  # Retry up to 5 times if file is locked
                    try:
                        df = pd.read_excel("data.xlsx", sheet_name="marketwatch", dtype=str)
                        df = df.where(pd.notna(df), "")  # Convert NaN to empty strings
                        data_json = df.to_json(orient="records")
                        
                        await websocket.send(data_json)
                        print("✅ Sent updated data")

                        # Set B2 to 2 after finishing the data reading process
                        sheet.range("B2").value = 2
                        print("💬 Finished reading data.")
                        break  # Exit retry loop if successful
                    except PermissionError:
                        print(f"⚠️ Excel file locked (Attempt {attempt + 1}/5)")
                        await asyncio.sleep(1)
                        attempt += 1
            else:
                print("⏳ A2 is 0, waiting...")

            await asyncio.sleep(5)  # Wait before sending next update
    except websockets.exceptions.ConnectionClosed as e:
        print(f"❌ Client disconnected: {e}")

async def handler(websocket, path):
    """Handle WebSocket connection"""
    print("✅ Client connected")
    sheet = wb.sheets["marketwatch"]
    await send_data(websocket, sheet)

async def main():
    sheet = wb.sheets["marketwatch"]
    save_task = asyncio.create_task(save_excel(sheet))
    server = await websockets.serve(handler, "smig.ddns.net", 9001)  # Use DDNS hostname
    print("🚀 WebSocket Server Started on ws://smig.ddns.net:9001")
    await server.wait_closed()

asyncio.run(main())
