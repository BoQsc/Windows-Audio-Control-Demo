![release](https://github.com/user-attachments/assets/9dd4851b-f4b6-4190-8a53-b5a81773d920)

# Windows Audio Control Demo

This project demonstrates how to control Windows audio devices using Python, the Windows Core Audio APIs, and Tkinter. It leverages Python’s built-in `ctypes` module to interact with COM objects and provides a graphical interface for:

- Enumerating active audio endpoint devices (render devices).
- Displaying device friendly names.
- Controlling the master volume (volume up, volume down, mute/unmute, and slider control).
- Switching the default audio render endpoint using a reverse-engineered (undocumented) interface.

> **Note:** Switching the default audio device uses an undocumented API (IPolicyConfigVista) and may require Administrator privileges. Use at your own risk.

## Features

- **Device Enumeration:** Lists active render devices with friendly names.
- **Volume Control:** Increase, decrease, and mute/unmute volume.
- **Volume Slider:** Adjust the master volume level (0–100%).
- **Default Device Switching:** Change the default audio endpoint using IPolicyConfigVista.
- **Tkinter GUI:** An intuitive graphical interface for interacting with audio controls.

## Requirements

- **Operating System:** Windows Vista/7 or later
- **Python Version:** 3.6 or later
- ~~**Administrator Privileges:** Required for switching the default audio endpoint.~~
- **Tkinter:** Typically included with Python on Windows.
