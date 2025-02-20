import ctypes
from ctypes import POINTER, byref, c_void_p, c_ulong

# Constants
COINIT_MULTITHREADED = 0x0
CLSCTX_ALL = 23  # Using CLSCTX_ALL for broad compatibility

# Load ole32.dll functions
ole32 = ctypes.windll.ole32

# Define the GUID structure as per Windows definitions
class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8)
    ]

def create_guid(guid_str):
    """
    Converts a GUID string (e.g., "BCDE0395-E52F-467C-8E3D-C4579291692E")
    into a GUID structure.
    """
    parts = guid_str.split('-')
    data1 = int(parts[0], 16)
    data2 = int(parts[1], 16)
    data3 = int(parts[2], 16)
    # Concatenate the 4th and 5th parts (which form 8 bytes)
    data4_bytes = bytes.fromhex(parts[3] + parts[4])
    data4 = (ctypes.c_ubyte * 8).from_buffer_copy(data4_bytes)
    return GUID(data1, data2, data3, data4)

# GUIDs extracted from the headers:
# CLSID for MMDeviceEnumerator
CLSID_MMDeviceEnumerator = create_guid("BCDE0395-E52F-467C-8E3D-C4579291692E")
# IID for IMMDeviceEnumerator interface
IID_IMMDeviceEnumerator = create_guid("A95664D2-9614-4F35-A746-DE8DB63617E6")
# IID for IAudioEndpointVolume interface (for later use)
IID_IAudioEndpointVolume = create_guid("5CDF2C82-841E-4546-9722-0CF74078229A")

def init_com():
    """Initialize the COM library."""
    hr = ole32.CoInitializeEx(None, COINIT_MULTITHREADED)
    if hr < 0:
        raise ctypes.WinError(hr)
    print("COM initialized successfully.")

def create_device_enumerator():
    """Creates an instance of MMDeviceEnumerator and returns its pointer."""
    pEnumerator = c_void_p()
    hr = ole32.CoCreateInstance(
        byref(CLSID_MMDeviceEnumerator),
        None,
        CLSCTX_ALL,
        byref(IID_IMMDeviceEnumerator),
        byref(pEnumerator)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    print("IMMDeviceEnumerator created successfully.")
    return pEnumerator

if __name__ == "__main__":
    # Step 1: Initialize COM
    init_com()
    
    # Step 2: Create the IMMDeviceEnumerator object
    enumerator = create_device_enumerator()
    
    # At this point, 'enumerator' holds a pointer to the IMMDeviceEnumerator COM object.
    # Further steps will involve using this pointer to enumerate audio devices,
    # get the default endpoint, and eventually activate IAudioEndpointVolume.
