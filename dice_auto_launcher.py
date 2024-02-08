import subprocess
import threading
from pynput import keyboard
import configparser
import os

starting_script_directory = os.path.dirname(os.path.abspath(__file__))
settings_file_path = os.path.join(starting_script_directory, '/home/gee/Downloads/ubuntu/settings.ini')

# Update the working_directory value in the settings.ini file
with open(settings_file_path, 'r') as f:
    lines = f.readlines()

with open(settings_file_path, 'a') as f:
    for line in lines:
        if line.startswith('working_directory = '):
            line = (f"working_directory = {starting_script_directory}")
            f.write(line)

# Read the settings values from the settings file
config = configparser.ConfigParser()

config.read('/home/gee/Downloads/ubuntu/settings.ini')
runs = int(config.get('cycles', 'runs'))
directory = config.get('home', 'working_directory')

def repeat_process(n):
    for _ in range(n):
        process = subprocess.Popen(['python3', f"{directory}/vm_chrome_linux.py"])
        # process = subprocess.Popen(['python3', 'create_tabs_auto.py'])
        process.wait()  # Wait for the process to finish before repeating
        if not running:
            exit()


running = True

def stop_script():
    print("Stopping script...")
    global running
    running = False



# Create a separate thread for the process
process_thread = threading.Thread(target=repeat_process, args=(runs,))
process_thread.start()


counter = 0
while running:
    counter += 1
    if counter < 2:
        # Create a listener for the escape key press event
        listener = keyboard.Listener(on_press=lambda key: stop_script() if key == keyboard.Key.esc else None)
        listener.start()
    if not listener.is_alive():
        stop_script()
        exit()
        break

# Terminate the process thread abruptly
process_thread._tstate_lock.release()
process_thread._stop()

