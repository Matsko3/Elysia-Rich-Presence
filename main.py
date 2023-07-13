import time
import pygetwindow as gw
import json
import configparser
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
    print("\x1b[38;5;91mWaiting for StarRail.exe to be running...\x1b[38;5;91m")
    time.sleep(5)
    if not gw.getWindowsWithTitle("StarRail"):
        print("StarRail.exe is not running.")

# Get the process ID of StarRail.exe
process_id = subprocess.run(['tasklist', '/fi', 'IMAGENAME eq StarRail.exe', '/fo', 'csv'], capture_output=True, text=True).stdout
process_lines = process_id.splitlines()
if len(process_lines) > 1:
    process_id = process_lines[1].split(',')[1].strip('"')
    print("\x1b[38;5;91mStarRail.exe found with process ID:\x1b[38;5;212m", process_id)
    print("\x1b[38;5;212mStarting the Rich Presence\x1b[38;5;212m")
    print("                                ╥▄")
    print("                               ╓██▌")
    print("                               █▓▓█")
    print("                             ╓█▓▓▓█")
    print("                            ▄█▓█▓▓█")
    print("                         ▄██████▓█▀")
    print("          ▀██    ▄╖   ┌██████████▌ ╓╖╥╦╦╦╦╦▄╗╦╦╦╗╖╖╓")
    print("           ▀██▄  ▐▓▓ ▄██████████▀╜▒╓╥▄▄▓ ░░▒▒▒╙  ╙░╙╜▀╬╥")
    print("            ▀██─ █▒╢█████████▀▀▀▒╜▀▓▓█▀    ╙          ╙▒▒▓╗")
    print("             ███╒▓▒╢████████▄▒▄▓▓███▀▄▓▓▓▓╖   ╓╬▓▓╗╖     ╙▒╙╗")
    print("             ▐███▒▒▓██▀╙╙▀▀████▀▀░▄▓▓▓╜╜╜╜    ╙╜╨▓▓▓▓▓╖      ╙╦")
    print("             ▐██▒║▒██▌╦╬▓▓██▓▄▄┐╩╜   ░             ░ ╙▓▓╗      ╙╗")
    print("             ▐██░▒▒█▀▒▒▒▒▒▒▄████████▄▄╓         ╙╗      ╫▓╖     └▓")
    print("             ▐██▌▄▓▒▒▒▓▓██████████████▀░░        ╙╣  ░    ▓╗      ▓")
    print("              ╙███░▒▓▓████████████▀┌   ░░░        ║┐ ╙╢    ╙░    ░ ▓")
    print("                 █▄║▓██▀╢▒▀▀▀▀░╟╣     ░      ╗    ╠▓▄  ╟╗        ░░ ╙╦       ╓")
    print("                  ████▒╜  ╓╣▒ ▓▌   ░         ▓   ░╠▓▒▓  ║▒╖  ░    ╙╖   ╙╨╤═╦▓")
    print("                 ╟███╜░   ╫▒░╟▓▒  ░░        ┌▓  ▒░╟▌ ╟▓  ╢▒▒ ░    ░░╙▓╥╥╥╣╜")
    print("                ▐▌▐█▌└░  ░▓▒░▓▓▒  ░░   ░ ░ ░╠▓░ ╢▒▓▒  ╠▌ ╟▒▒▒▒░      ╙╗")
    print("               ╔▓╢╣╜░░░   ▓▒╫╣▓▒ ░░ ░░▒░ ░ ▒▓▓░╓▒╟▓    ▓╖╠▓▓▒╫░ ░░  ▒ ╫")
    print("              ╒▓▒╢▓░░░░░░ ╟╣▓▒▓▒ ░░  ╟▒ ░ ▒▒▓╢▒╢╫▓     ╠▓╟▓╟╣╣  ░░ ░╫▒ ▓")
    print("              ▓╫╩▀▓╣░░░░░░▒╢▓▒╙▓▓▓▓▓▓▓╣╬▓▓▓▀╙╨╫▓╜  ╥▓▓▓▓╣╢░▒▓▒░░▒░░▒▓╝╝╨▀┐")
    print("             ╟ ╙▓╖ ╟▒▒░░░▒▒▒▒▒ ╠╬╬╬╬╬╬            ║▓╬╬╬▓▓  ╟╣░░▒░ ▒╫▒ ╓╦╜╗")
    print("             ╬ ╟▓▓▓▓▓▒▒╖░░╙▒╣░  ░░░░░░             ░╓╥╝▒  ╓╣╖▒▒░░▒▒▓▓╝▓▒░╙╖")
    print("            ╓▒╓▌▓╢▒╙╫▓▒▒▒▒▒╖▒╫╗  ▒░░░    ELYSIA    ░░▒▒  ╓▒▒╣▒╥▒▒▒▓╜░╓▓▒  ▓")
    print("            ╟▒▓  ╚▓▒░▒▓▓▒▒▒▒▒▒▒▓╖                      ╓▓▓▒▒▒▒▒╫▓░░░░▓ ╟ ░▓")
    print("             ▒╢    ╟╖▒░░░╙╨▓▓▒▒▒▒╨╨╜      ╥╖╖╖─    ╨▓▓▒▒▒▒▒▄▓▓╜░░░░╓▓  ╣ ▓")
    print("             ╟▒     ╙╣▒▒▒░▒▒╙▓▓▓▓▓▄╖             ╖  ╓▄▓▓▓▓▓▀░░░░▒░╬╜  ╥╬▓")
    print("              ╫╖      ╫╣▓╖░▒░░░▒╙▓▓▓▓▓▓▄▄▄╥╦╥╦▄▄▄▓▒▀▀▀╜╙░░▒░░░▒░╫▓▓▓▓▓▓")
    print("               ╙╬╖      ▓▓▓╩╬▒░▒▒▒╥╣▓█████▓███▓███▓▓╖░░▒▒▒░▒▒▓▓▀╙╙╙")
    print("               ╥  ╙    ╓╣▒╜▒▒▒╢▓▓╣▓██▓▓█████████████▓▓▒╜▓▒▒╙▒▒▒▓╗      ╓▒")
    print("                ╙▓▄╬▓▓▒▒▒╥╣╣╣╢╢╬╩▀╖███████▓███████▓▓█▌█▓▓╣▓╬╢╫╬╣╣╣▓▓▓▓▓╙")
    print("                    ╙╜╜╨╜╜╙        ▀███▓▓▀█▀█▀▀   ▀╩╩▀▀▌     ╙╙╜╨╜╜╙\x1b[38;5;212m")

else:
    print("\x1b[38;5;91mUnable to retrieve process ID for StarRail.exe\x1b[38;5;91m")
    sys.exit()

# Read the configs INI file
config = configparser.ConfigParser()
config.read('config/configs.ini')

# Read the fallback JSON file
with open('config/fallback.json', 'r') as file:
    refresh_data = json.load(file)

words_to_search = []
for section in config.sections():
    word_data = {}
    word_data['word'] = config.get(section, 'word').replace('\\n', '\n')
    word_data['details'] = config.get(section, 'details')
    word_data['state'] = config.get(section, 'state')
    word_data['large_image'] = config.get(section, 'large_image')
    word_data['small_image'] = config.get(section, 'small_image')
    words_to_search.append(word_data)

history = []

# handling exceptions and restarting the script
def handle_exception():
    print("An exception occurred. Restarting the script...")
    time.sleep(5)  # Wait for a few seconds before restarting

while True:
    try:
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
            try:
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

                # Used whenever assets are max
                if large_image.startswith('https://'):
                    large_image = large_image  # Use the provided URL for large image

                if small_image.startswith('https://'):
                    small_image = small_image  # Use the provided URL for small image

                # Updating the Rich Presence activity
                RPC.update(
                    details=details,
                    state=state,
                    large_image=large_image,
                    small_image=small_image,
                    start=start_time
                )

                # Printing Rich Presence data
                print("\x1b[38;5;91mRich Presence Updated:\x1b[38;5;91m", "Details:", details, "State:", state, "Large Image:", large_image, "Small Image:", small_image)
                # Sleep for some time before capturing the next text
                time.sleep(3)

            except Exception as e:
                handle_exception()

    except Exception as e:
        handle_exception()

    finally:
        # Close the Rich Presence connection before restarting
        if 'RPC' in locals() and RPC is not None:
            RPC.close()
