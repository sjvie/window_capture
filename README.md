# Window Capture Application

This application captures information about open windows on a Windows system at regular intervals. It runs in the background with a system tray icon and saves the captured data to CSV files.

## Disclaimer

**This code was generated using a large language model (LLM). I take no responsibility for the code quality, security vulnerabilities, or any other issues that may arise from using this code. Use at your own risk.**

## Features

- Captures information about all visible windows, including:
  - Window title
  - Process name
  - Active (foreground) status
  - Window order (z-order)
- Runs silently in the system tray
- Captures window information every minute
- Creates a new CSV file for each session
- Logs operations and errors for easy troubleshooting

## Requirements

- Windows operating system
- Python 3.6+
- Required Python packages (see `requirements.txt` for versions):
  - pywin32
  - psutil
  - schedule
  - pystray
  - Pillow

## Installation

1. Clone this repository or download the source code.
2. Create a virtual environment:
   ```
   python -m venv window_capture_env
   ```
3. Activate the virtual environment:
   ```
   window_capture_env\Scripts\activate
   ```
4. Install the required packages using the `requirements.txt` file:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Ensure your virtual environment is activated.
2. Run the `run_window_capture.bat` file to start the application.
3. The application will run in the background with an icon in the system tray.
4. To stop the application, right-click the system tray icon and select "Quit".

## Setting Up Autostart

To make the Window Capture Application start automatically when you log into Windows:

1. Right-click on `run_window_capture.bat` and select "Create shortcut".
2. Press `Win + R` to open the Run dialog.
3. Type `shell:startup` and press Enter. This opens the Startup folder.
4. Move the shortcut you created in step 1 into this Startup folder.

Now, the application will start automatically each time you log into your Windows account.

## File Structure

- `window_capture.py`: The main Python script that captures window information.
- `run_window_capture.bat`: A batch file to run the script with the correct Python environment.
- `requirements.txt`: A list of Python package dependencies.
- `captures/`: Directory where CSV files with captured data are stored.
- `logs/`: Directory containing the log file (`window_capture.log`).

## CSV File Format

Each capture session creates a new CSV file with the following columns:
- `id`: Timestamp of the capture (Unix timestamp)
- `timestamp`: ISO format timestamp of when each window was captured
- `title`: Window title
- `process`: Name of the process that owns the window
- `is_active`: Boolean indicating if the window was active (in the foreground)
- `order`: Z-order of the window (1 being topmost)

## Logging

The application logs its operations and any errors to `logs/window_capture.log`. Check this file if you encounter any issues or want to monitor the application's activity.
