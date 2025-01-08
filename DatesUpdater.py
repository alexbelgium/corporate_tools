import os
import time
from datetime import datetime, timedelta

# Set the max age in days and the path to the current folder
max_age = 200
mypath = os.path.abspath(os.path.dirname(__file__))

print(
    r"""\
     _______ __                             __      __                    ___
    / ____(_) /__  _____   __  ______  ____/ /___ _/ /____  _____   _   _<  /
   / /_  / / / _ \/ ___/  / / / / __ \/ __  / __ `/ __/ _ \/ ___/  | | / / / 
  / __/ / / /  __(__  )  / /_/ / /_/ / /_/ / /_/ / /_/  __/ /      | |/ / /  
 /_/   /_/_/\___/____/   \__,_/ .___/\__,_/\__,_/\__/\___/_/       |___/_/   
                             /_/                                             
"""
)

print(f"Starting the script to update files and folders older than {max_age} days in the current folder and its subfolders.")
print(f"The folder in which we will look is {mypath}\n")

# Progress bar function
def progress_bar(current, total, message=""):
    bar_length = 40
    percent = float(current) / total
    arrow = "=" * int(round(percent * bar_length))
    spaces = " " * (bar_length - len(arrow))
    print(f"\r{message} [{arrow}{spaces}] {int(percent * 100)}%", end="", flush=True)

# Count files and folders older than max_age days
print(f"Counting the number of files and folders older than {max_age} days...")
files_to_update = []
current_time = time.time()
age_limit = current_time - (max_age * 86400)

all_items = []
for root, dirs, files in os.walk(mypath):
    all_items.extend([os.path.join(root, d) for d in dirs])
    all_items.extend([os.path.join(root, f) for f in files])

total_items = len(all_items)
for i, item_path in enumerate(all_items, start=1):
    item_mtime = os.path.getmtime(item_path)
    if item_mtime < age_limit:
        files_to_update.append(item_path)
    progress_bar(i, total_items, message="Counting files and folders")

print("\nDone.")
print(f"Number of files and folders to update timestamps: {len(files_to_update)}")
if len(files_to_update) > 0:
    input("Press Enter to start updating timestamps...")

# Update file and folder timestamps
for i, item_path in enumerate(files_to_update, start=1):
    os.utime(item_path, None)  # Update timestamp to current time
    progress_bar(i, len(files_to_update), message="Updating timestamps")

print("\nTimestamp updating complete.\n")
input("Press Enter to exit...")
