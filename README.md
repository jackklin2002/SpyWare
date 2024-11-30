# SpyWare
Keylogger and Utility Toolkit
This repository hosts a versatile application designed to provide multiple functionalities, including keylogging, browser history retrieval, Wi-Fi information scanning, port scanning, and system information logging, all accessible through an intuitive graphical user interface (GUI).

The keylogger feature captures keystrokes in real-time and saves them to a log file, allowing you to start and stop the logging process directly from the GUI. The application also retrieves browser history from popular browsers like Google Chrome, Mozilla Firefox, Microsoft Edge, and Safari, exporting the data to Excel files for easy analysis.

Additionally, the Wi-Fi information scanner displays network details such as IP address, MAC address, default gateway, and DNS servers (if available). The port scanning tool checks open ports on the local machine within a specified range (default: ports 1–1024) and exports the results to an Excel file.

The system information logger records essential details about the machine, including the current date, IP address, processor, operating system, and host name. This data is saved in an Excel file for reference. To enhance monitoring, the toolkit also includes a screenshot capture feature, saving images with timestamps to the screenshots directory.

The application is built with Python, leveraging libraries such as tkinter for the GUI, pandas for data manipulation, and pynput for keylogging functionality. Installation is straightforward: clone the repository, install the required dependencies using the provided requirements.txt file, and run the app.py script to launch the application.

Users can interact with the application via the GUI to perform tasks such as starting the keylogger, retrieving browser history, scanning for open ports, viewing Wi-Fi network details, or logging system information. While the toolkit is feature-rich, it should be used responsibly. Tools like keyloggers and port scanners can have privacy implications, so ensure you have explicit permission to use them. Misuse may lead to legal consequences, depending on your jurisdiction.
