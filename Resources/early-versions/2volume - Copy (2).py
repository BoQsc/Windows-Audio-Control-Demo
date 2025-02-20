import ctypes
from ctypes import POINTER, byref, c_void_p, c_ulong, c_int, c_wchar_p
import tkinter as tk
from tkinter import ttk

# ============================================================
# Constants and Definitions
# ============================================================

# COM initialization and context constants
COINIT_MULTITHREADED = 0x0
CLSCTX_ALL = 23

# Enumerations for audio endpoint selection (from mmdeviceapi.h)
EDataFlow_eRender = 0   # Rendering devices (e.g., speakers)
ERole_eConsole    = 0   # Console role

# Load COM library functions
ole32 = ctypes.windll.ole32

# -------------------------------
# GUID Structure and Helper
# -------------------------------
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
    data4_bytes = bytes.fromhex(parts[3] + parts[4])
    data4 = (ctypes.c_ubyte * 8).from_buffer_copy(data4_bytes)
    return GUID(data1, data2, data3, data4)

# GUIDs from header files:
CLSID_MMDeviceEnumerator = create_guid("BCDE0395-E52F-467C-8E3D-C4579291692E")
IID_IMMDeviceEnumerator   = create_guid("A95664D2-9614-4F35-A746-DE8DB63617E6")
IID_IAudioEndpointVolume  = create_guid("5CDF2C82-841E-4546-9722-0CF74078229A")

# ============================================================
# COM Initialization and IMMDeviceEnumerator Creation
# ============================================================

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

# ============================================================
# IMMDeviceEnumerator Interface Definition (VTable)
# ============================================================
class IMMDeviceEnumeratorVTable(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
        ("AddRef",         ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("Release",        ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("EnumAudioEndpoints", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_int, c_ulong, POINTER(c_void_p))),
        ("GetDefaultAudioEndpoint", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_int, c_int, POINTER(c_void_p))),
        ("GetDevice",      ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(c_void_p))),
        ("RegisterEndpointNotificationCallback", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("UnregisterEndpointNotificationCallback", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p))
    ]

class IMMDeviceEnumerator_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IMMDeviceEnumeratorVTable))]

def get_default_endpoint(enumerator):
    """
    Uses the IMMDeviceEnumerator interface to obtain the default audio endpoint.
    """
    enumerator_iface = ctypes.cast(enumerator, POINTER(IMMDeviceEnumerator_Interface))
    default_endpoint = c_void_p()
    hr = enumerator_iface.contents.lpVtbl.contents.GetDefaultAudioEndpoint(
        enumerator_iface,      # 'this' pointer
        EDataFlow_eRender,     # render devices
        ERole_eConsole,        # console role
        byref(default_endpoint)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Default audio endpoint obtained:", default_endpoint)
    return default_endpoint

# ============================================================
# IMMDevice Interface Definition (for Activate)
# ============================================================
class IMMDeviceVTable(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
        ("AddRef",         ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("Release",        ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("Activate",       ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), c_ulong, c_void_p, POINTER(c_void_p))),
        ("OpenPropertyStore", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_ulong, POINTER(c_void_p))),
        ("GetId",          ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_wchar_p))),
        ("GetState",       ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_ulong)))
    ]

class IMMDevice_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IMMDeviceVTable))]

def activate_audio_endpoint_volume(endpoint):
    """
    Activates the IAudioEndpointVolume interface for the given endpoint.
    """
    endpoint_iface = ctypes.cast(endpoint, POINTER(IMMDevice_Interface))
    audio_endpoint_volume = c_void_p()
    hr = endpoint_iface.contents.lpVtbl.contents.Activate(
        endpoint_iface,
        byref(IID_IAudioEndpointVolume),  # request IAudioEndpointVolume
        CLSCTX_ALL,
        None,
        byref(audio_endpoint_volume)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    print("IAudioEndpointVolume activated successfully:", audio_endpoint_volume)
    return audio_endpoint_volume

# ============================================================
# IAudioEndpointVolume Interface Definition (VTable)
# ============================================================
class IAudioEndpointVolumeVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
        ("AddRef",         ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("Release",        ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("RegisterControlChangeNotify", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("UnregisterControlChangeNotify", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("GetChannelCount", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_uint))),
        ("SetMasterVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_float, c_void_p)),
        ("SetMasterVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_float, c_void_p)),
        ("GetMasterVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_float))),
        ("GetMasterVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_float))),
        ("SetChannelVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_uint, ctypes.c_float, c_void_p)),
        ("SetChannelVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_uint, ctypes.c_float, c_void_p)),
        ("GetChannelVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_uint, POINTER(ctypes.c_float))),
        ("GetChannelVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_uint, POINTER(ctypes.c_float))),
        ("SetMute", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_int, c_void_p)),
        ("GetMute", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_int))),
        ("GetVolumeStepInfo", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_uint), POINTER(ctypes.c_uint))),
        ("VolumeStepUp", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("VolumeStepDown", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("QueryHardwareSupport", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_ulong))),
        ("GetVolumeRange", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_float), POINTER(ctypes.c_float), POINTER(ctypes.c_float)))
    ]

class IAudioEndpointVolume_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IAudioEndpointVolumeVtbl))]

# ============================================================
# Volume Control Helper Functions
# ============================================================
def volume_step_up(audio_volume):
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    hr = volume_iface.contents.lpVtbl.contents.VolumeStepUp(volume_iface, None)
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Volume stepped up.")

def volume_step_down(audio_volume):
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    hr = volume_iface.contents.lpVtbl.contents.VolumeStepDown(volume_iface, None)
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Volume stepped down.")

def set_mute(audio_volume, mute):
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    hr = volume_iface.contents.lpVtbl.contents.SetMute(volume_iface, int(mute), None)
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Mute set to", mute)

def get_mute(audio_volume):
    mute_val = c_int()
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    hr = volume_iface.contents.lpVtbl.contents.GetMute(volume_iface, byref(mute_val))
    if hr < 0:
        raise ctypes.WinError(hr)
    return bool(mute_val.value)

# ============================================================
# Tkinter GUI for Volume Control
# ============================================================
class VolumeControlApp(tk.Tk):
    def __init__(self, audio_volume):
        super().__init__()
        self.title("Volume Control Demo")
        self.audio_volume = audio_volume

        # Create buttons for actions
        self.btn_up = ttk.Button(self, text="Volume Up", command=self.volume_up)
        self.btn_down = ttk.Button(self, text="Volume Down", command=self.volume_down)
        self.btn_toggle = ttk.Button(self, text="Toggle Mute", command=self.toggle_mute)
        self.lbl_status = ttk.Label(self, text="")

        # Layout the widgets
        self.btn_up.pack(pady=5)
        self.btn_down.pack(pady=5)
        self.btn_toggle.pack(pady=5)
        self.lbl_status.pack(pady=5)

        self.update_status()

    def volume_up(self):
        volume_step_up(self.audio_volume)
        self.update_status()

    def volume_down(self):
        volume_step_down(self.audio_volume)
        self.update_status()

    def toggle_mute(self):
        current = get_mute(self.audio_volume)
        set_mute(self.audio_volume, not current)
        self.update_status()

    def update_status(self):
        muted = get_mute(self.audio_volume)
        status = "Muted" if muted else "Unmuted"
        self.lbl_status.config(text="Current Status: " + status)

# ============================================================
# Main Function: Execute the Steps and Launch GUI
# ============================================================
def main():
    # Step 1: Initialize COM
    init_com()
    
    # Step 2: Create the IMMDeviceEnumerator object
    enumerator = create_device_enumerator()
    
    # Step 3: Get the default audio endpoint (e.g., speakers)
    default_endpoint = get_default_endpoint(enumerator)
    
    # Step 4: Activate the IAudioEndpointVolume interface for volume control
    audio_volume = activate_audio_endpoint_volume(default_endpoint)
    
    # Launch Tkinter GUI for volume control
    app = VolumeControlApp(audio_volume)
    app.mainloop()

if __name__ == "__main__":
    main()
