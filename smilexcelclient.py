import asyncio
import websockets
import json
import win32com.client
import time

async def fetch_data():
    """Fetch data from WebSocket server"""
    while True:
        try:
            async with websockets.connect("ws://49.36.65.63:9001") as websocket:  # Replace with your WebSocket URL  192.168.29.105:9001
                print("✅ Connected to WebSocket server")

                while True:
                    try:
                        data = await websocket.recv()  # Receive JSON data
                        print("✅ Received data:", data)
                        update_excel(data)  # Update Excel sheet
                    except websockets.exceptions.ConnectionClosedError as e:
                        print(f"❌ Connection Closed: {e}")
                        break  # Restart connection on next loop
                    except Exception as e:
                        print(f"❌ Unexpected error: {e}")
                    await asyncio.sleep(6)  # Wait before fetching again

        except ConnectionRefusedError:
            print("❌ ERROR: Could not connect to the WebSocket server!")
        except Exception as e:
            print(f"❌ Client error: {e}")
        await asyncio.sleep(7)  # Retry connection every 7 seconds

def update_excel(json_data):
    try:
        excel = win32com.client.Dispatch("Excel.Application")  # Open Excel
        excel.Visible = True  # Keep Excel visible

        workbook = None
        for wb in excel.Workbooks:
            if wb.Name == "client_data.xlsx":
                workbook = wb
                break

        if workbook is None:
            print("❌ ERROR: client_data.xlsx is not open!")
            return  # Exit if file is not open

        sheet = workbook.Sheets("marketwatch")

        # Check if A2 contains 1, if not, wait for 1 second before checking again
        while True:
            a2_value = sheet.Cells(2, 1).Value  # A2 corresponds to row=2, column=1
            if a2_value == 1:
                break
            else:
                print("❌ A2 is not 1, waiting for 1 second...")
                time.sleep(1)  # Wait 1 second before checking again

        # Load JSON data
        data = json.loads(json_data)

        # Write data to Excel starting from A3
        for i, row in enumerate(data, start=3):  # Start at row 3
            for j, (key, value) in enumerate(row.items(), start=1):  # Columns start at 1 (A)
                sheet.Cells(i, j).Value = value

        print("✅ Data successfully written to Excel starting from A3!")

        # Save the workbook after updating
        workbook.Save()
        print("✅ Excel workbook saved!")

    except Exception as e:
        print(f"❌ Error updating Excel: {e}")

if __name__ == "__main__":
    asyncio.run(fetch_data())
