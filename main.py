import time
import pygetwindow as gw
import json
from pypresence import Presence
from PIL import ImageGrab
from ocr import extract_text_from_image
import subprocess
import sys

# Check if StarRail.exe is running
def is_star_rail_running():
    process_list = subprocess.run(['tasklist'], capture_output=True, text=True).stdout
    return 'StarRail.exe' in process_list

# Wait for StarRail.exe launch
while not is_star_rail_running():
    print("Waiting for StarRail.exe to be running...")
    time.sleep(5)
    if not gw.getWindowsWithTitle("StarRail"):
        print("StarRail.exe is not running.")

# Get the process ID of StarRail.exe
process_id = subprocess.run(['tasklist', '/fi', 'IMAGENAME eq StarRail.exe', '/fo', 'csv'], capture_output=True, text=True).stdout
process_lines = process_id.splitlines()
if len(process_lines) > 1:
    process_id = process_lines[1].split(',')[1].strip('"')
    print("StarRail.exe found with process ID:", process_id)
    print("Starting the Rich Presence")
else:
    print("Unable to retrieve process ID for StarRail.exe")
    sys.exit()

# Read the config JSON file
with open('config/configs.json', 'r') as file:
    rpc_data = json.load(file)
    words_to_search = rpc_data['words']

# Read the fallback JSON file
with open('config/fallback.json', 'r') as file:
    refresh_data = json.load(file)

# Initialize/Start the Discord RPC client
RPC = Presence('724484541825286205')

RPC.connect()

# Auto refreshing data
refresh_details = refresh_data['refresh_details']
refresh_state = refresh_data['refresh_state']
refresh_large_image = refresh_data['refresh_large_image']
refresh_small_image = refresh_data['refresh_small_image']

start_time = int(time.time())

# Main loop
while True:
    # Capture the screen text
    screen_image = ImageGrab.grab()

    # Use OCR to extract text from the image
    extracted_text = extract_text_from_image(screen_image)

    # Initialize variables for Rich Presence data
    details = refresh_details
    state = refresh_state
    large_image = refresh_large_image
    small_image = refresh_small_image

    # Check if word matches the extracted text
    for word_data in words_to_search:
        word_lines = word_data['word'].split('\n')
        if all(line in extracted_text for line in word_lines):
            details = word_data['details']
            state = word_data['state']
            large_image = word_data['large_image']
            small_image = word_data['small_image']
            break  # Exit the loop once a match is found

    # Updating the Rich Presence activity
    RPC.update(
        details=details,
        state=state,
        large_image=large_image,
        small_image=small_image,
        start=start_time
    )

    # Printing Rich Presence data
    print("Rich Presence Updated:")
    print("Details:", details)
    print("State:", state)
    print("Large Image:", large_image)
    print("Small Image:", small_image)
    print("-----------------------")

    # Sleep for some time before capturing the next text
    time.sleep(4)

RPC.close()
