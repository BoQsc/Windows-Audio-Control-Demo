import ctypes
from ctypes import POINTER, byref, c_void_p, c_ulong, c_int, c_wchar_p, c_uint
import tkinter as tk
from tkinter import ttk

# ============================================================
# Constants and Definitions
# ============================================================

# COM and context constants
COINIT_MULTITHREADED = 0x0
CLSCTX_ALL = 23

# Device selection enumerations (from mmdeviceapi.h)
EDataFlow_eRender = 0   # Render devices (e.g., speakers)
ERole_eConsole    = 0   # Console role
ERole_eMultimedia = 1   # Multimedia role
ERole_eCommunications = 2  # Communications role

# Device state mask
DEVICE_STATE_ACTIVE = 0x00000001

# Load COM functions from ole32.dll
ole32 = ctypes.windll.ole32

# -------------------------------
# GUID and Helper Structures
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

# For switching default device, we use the undocumented IPolicyConfig interface.
# These GUIDs are commonly used in the community.
CLSID_CPolicyConfigClient = create_guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")
IID_IPolicyConfig           = create_guid("F8679F50-850A-41CF-9C72-430F290290C8")

# -------------------------------
# PROPERTYKEY and PROPVARIANT (for friendly names)
# -------------------------------
class PROPERTYKEY(ctypes.Structure):
    _fields_ = [
        ("fmtid", GUID),
        ("pid", ctypes.c_ulong)
    ]

# For our purposes we only handle VT_LPWSTR (VT value 31)
VT_LPWSTR = 31

class PROPVARIANT(ctypes.Structure):
    _fields_ = [
        ("vt", ctypes.c_ushort),
        ("wReserved1", ctypes.c_ushort),
        ("wReserved2", ctypes.c_ushort),
        ("wReserved3", ctypes.c_ushort),
        ("pwszVal", ctypes.c_wchar_p)
    ]

# Define PKEY_Device_FriendlyName:
# {A45C254E-DF1C-4EFD-8020-67D146A850E0}, 14
PKEY_Device_FriendlyName = PROPERTYKEY(create_guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14)

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
# IMMDeviceEnumerator Interface (VTable and Interface)
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
        enumerator_iface,
        EDataFlow_eRender,
        ERole_eConsole,
        byref(default_endpoint)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Default audio endpoint obtained:", default_endpoint)
    return default_endpoint

# ============================================================
# IMMDevice Interface (for Activate, OpenPropertyStore, and GetId)
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
        byref(IID_IAudioEndpointVolume),
        CLSCTX_ALL,
        None,
        byref(audio_endpoint_volume)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    print("IAudioEndpointVolume activated successfully:", audio_endpoint_volume)
    return audio_endpoint_volume

def get_device_id(device):
    """
    Uses the IMMDevice interface to get the device's ID (a string).
    """
    device_iface = ctypes.cast(device, POINTER(IMMDevice_Interface))
    pDeviceId = c_wchar_p()
    hr = device_iface.contents.lpVtbl.contents.GetId(device_iface, byref(pDeviceId))
    if hr < 0:
        raise ctypes.WinError(hr)
    return pDeviceId.value

# ============================================================
# IAudioEndpointVolume Interface (VTable and Interface)
# ============================================================
class IAudioEndpointVolumeVtbl(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
        ("AddRef",         ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("Release",        ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
        ("RegisterControlChangeNotify", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("UnregisterControlChangeNotify", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("GetChannelCount", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_uint))),
        ("SetMasterVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_float, c_void_p)),
        ("SetMasterVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, ctypes.c_float, c_void_p)),
        ("GetMasterVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_float))),
        ("GetMasterVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(ctypes.c_float))),
        ("SetChannelVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, ctypes.c_float, c_void_p)),
        ("SetChannelVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, ctypes.c_float, c_void_p)),
        ("GetChannelVolumeLevel", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, POINTER(ctypes.c_float))),
        ("GetChannelVolumeLevelScalar", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, POINTER(ctypes.c_float))),
        ("SetMute", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_int, c_void_p)),
        ("GetMute", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_int))),
        ("GetVolumeStepInfo", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_uint), POINTER(c_uint))),
        ("VolumeStepUp", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("VolumeStepDown", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_void_p)),
        ("QueryHardwareSupport", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_ulong))),
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

def get_master_volume(audio_volume):
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    level = ctypes.c_float()
    hr = volume_iface.contents.lpVtbl.contents.GetMasterVolumeLevelScalar(volume_iface, byref(level))
    if hr < 0:
        raise ctypes.WinError(hr)
    return level.value

def set_master_volume(audio_volume, value):
    volume_iface = ctypes.cast(audio_volume, POINTER(IAudioEndpointVolume_Interface))
    hr = volume_iface.contents.lpVtbl.contents.SetMasterVolumeLevelScalar(volume_iface, ctypes.c_float(value), None)
    if hr < 0:
        raise ctypes.WinError(hr)
    print("Master volume set to", value)

# ============================================================
# IPropertyStore Interface (for Friendly Names)
# ============================================================
class IPropertyStoreVtbl(ctypes.Structure):
    _fields_ = [
         ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
         ("AddRef", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("Release", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("GetCount", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_uint))),
         ("GetAt", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, POINTER(PROPERTYKEY))),
         ("GetValue", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(PROPERTYKEY), POINTER(PROPVARIANT)))
         # (SetValue and Commit omitted for brevity)
    ]

class IPropertyStore_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IPropertyStoreVtbl))]

def get_device_friendly_name(device):
    """
    Opens the property store for the device and retrieves the friendly name.
    """
    device_iface = ctypes.cast(device, POINTER(IMMDevice_Interface))
    pPropertyStore = c_void_p()
    hr = device_iface.contents.lpVtbl.contents.OpenPropertyStore(device_iface, 0, byref(pPropertyStore))
    if hr < 0:
        raise ctypes.WinError(hr)
    prop_store = ctypes.cast(pPropertyStore, POINTER(IPropertyStore_Interface))
    propvar = PROPVARIANT()
    hr = prop_store.contents.lpVtbl.contents.GetValue(prop_store, byref(PKEY_Device_FriendlyName), byref(propvar))
    if hr < 0:
        raise ctypes.WinError(hr)
    return propvar.pwszVal

# ============================================================
# IMMDeviceCollection Interface (for Enumerating Devices)
# ============================================================
class IMMDeviceCollectionVTable(ctypes.Structure):
    _fields_ = [
         ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
         ("AddRef", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("Release", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("GetCount", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(c_uint))),
         ("Item", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_uint, POINTER(c_void_p)))
    ]

class IMMDeviceCollection_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IMMDeviceCollectionVTable))]

def enumerate_audio_endpoints(enumerator):
    """
    Enumerates all active render devices and returns a list of tuples:
    (device pointer, friendly name)
    """
    enumerator_iface = ctypes.cast(enumerator, POINTER(IMMDeviceEnumerator_Interface))
    pCollection = c_void_p()
    hr = enumerator_iface.contents.lpVtbl.contents.EnumAudioEndpoints(
        enumerator_iface,
        EDataFlow_eRender,
        DEVICE_STATE_ACTIVE,
        byref(pCollection)
    )
    if hr < 0:
        raise ctypes.WinError(hr)
    collection_iface = ctypes.cast(pCollection, POINTER(IMMDeviceCollection_Interface))
    count = c_uint()
    hr = collection_iface.contents.lpVtbl.contents.GetCount(collection_iface, byref(count))
    if hr < 0:
        raise ctypes.WinError(hr)
    devices = []
    for i in range(count.value):
        pDevice = c_void_p()
        hr = collection_iface.contents.lpVtbl.contents.Item(collection_iface, i, byref(pDevice))
        if hr < 0:
            continue
        try:
            name = get_device_friendly_name(pDevice)
        except Exception as e:
            name = "Unknown Device"
        devices.append((pDevice, name))
    return devices

# ============================================================
# IPolicyConfig Interface (Undocumented, for switching default device)
# ============================================================
class IPolicyConfigVtbl(ctypes.Structure):
    _fields_ = [
         ("QueryInterface", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(GUID), POINTER(c_void_p))),
         ("AddRef", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("Release", ctypes.WINFUNCTYPE(ctypes.c_ulong, c_void_p)),
         ("GetMixFormat", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(c_void_p))),
         ("GetDeviceFormat", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_int, POINTER(c_void_p))),
         ("SetDeviceFormat", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_void_p, c_void_p)),
         ("GetProcessingPeriod", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_int, POINTER(ctypes.c_longlong), POINTER(ctypes.c_longlong))),
         ("SetProcessingPeriod", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(ctypes.c_longlong))),
         ("GetShareMode", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(c_void_p))),
         ("SetShareMode", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_void_p)),
         ("GetPropertyValue", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))),
         ("SetPropertyValue", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, POINTER(PROPERTYKEY), POINTER(PROPVARIANT))),
         ("SetDefaultEndpoint", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_int)),
         ("SetEndpointVisibility", ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, c_wchar_p, c_int))
    ]

class IPolicyConfig_Interface(ctypes.Structure):
    _fields_ = [("lpVtbl", POINTER(IPolicyConfigVtbl))]

# Define new GUIDs for IPolicyConfigVista based on the provided header snippet:
CLSID_CPolicyConfigVistaClient = create_guid("294935CE-F637-4E7C-A41B-AB255460B862")
IID_IPolicyConfigVista = create_guid("568B9108-44BF-40B4-9006-86AFE5B5A620")

def switch_default_device(device):
    """
    Switches the default audio endpoint to the given device.
    Uses the reverse‑engineered IPolicyConfigVista interface.
    Attempts to set the default endpoint for eConsole, eMultimedia, and eCommunications roles.
    If a call fails with "The tag is invalid" (error -2147023163), it logs a warning and continues.
    """
    device_id = get_device_id(device)
    print("Switching default device to ID:", device_id)

    # Create an instance of IPolicyConfigVista
    pPolicyConfig = c_void_p()
    hr = ole32.CoCreateInstance(
         byref(CLSID_CPolicyConfigVistaClient),
         None,
         CLSCTX_ALL,
         byref(IID_IPolicyConfigVista),
         byref(pPolicyConfig)
    )
    if hr < 0:
         raise ctypes.WinError(hr)
    print("IPolicyConfigVista activated successfully:", pPolicyConfig)

    # Cast to our IPolicyConfig interface (the methods are the same as in our previous definition)
    policy_config = ctypes.cast(pPolicyConfig, POINTER(IPolicyConfig_Interface))
    # Try to set the default endpoint for multiple roles.
    roles = [ERole_eConsole, ERole_eMultimedia, ERole_eCommunications]
    for role in roles:
        hr = policy_config.contents.lpVtbl.contents.SetDefaultEndpoint(policy_config, device_id, role)
        if hr < 0:
            # If error is "The tag is invalid" (-2147023163), log a warning and continue.
            if hr == -2147023163:
                print(f"SetDefaultEndpoint for role {role} failed with ERROR_INVALID_TAG, skipping.")
            else:
                raise ctypes.WinError(hr)
    print("Default device switched to:", device_id)

# ============================================================
# Tkinter GUI with Device Selection, Volume Control, Slider, and Default Switch
# ============================================================
class VolumeControlApp(tk.Tk):
    def __init__(self, enumerator, default_endpoint):
        super().__init__()
        self.title("Volume Control Demo with Device Switching")
        self.enumerator = enumerator

        # Enumerate devices
        self.devices = enumerate_audio_endpoints(self.enumerator)
        if not self.devices:
            raise Exception("No audio devices found.")
        
        # Determine default selection index
        self.default_index = 0
        for idx, (dev, name) in enumerate(self.devices):
            if dev.value == default_endpoint.value:
                self.default_index = idx
                break

        # Device selection dropdown
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(self, textvariable=self.device_var, state="readonly",
                                         values=[name for (_, name) in self.devices])
        self.device_combo.current(self.default_index)
        self.device_combo.bind("<<ComboboxSelected>>", self.on_device_selected)
        self.device_combo.pack(pady=5)

        # Button to set selected device as default
        self.btn_set_default = ttk.Button(self, text="Set as Default", command=self.set_as_default)
        self.btn_set_default.pack(pady=5)

        # Activate volume control for the initially selected device
        self.audio_volume = activate_audio_endpoint_volume(self.devices[self.default_index][0])

        # Buttons for volume control
        self.btn_up = ttk.Button(self, text="Volume Up", command=self.volume_up)
        self.btn_down = ttk.Button(self, text="Volume Down", command=self.volume_down)
        self.btn_toggle = ttk.Button(self, text="Toggle Mute", command=self.toggle_mute)
        self.lbl_status = ttk.Label(self, text="")

        self.btn_up.pack(pady=5)
        self.btn_down.pack(pady=5)
        self.btn_toggle.pack(pady=5)
        self.lbl_status.pack(pady=5)

        # Volume slider (0 to 100)
        self.volume_slider = tk.Scale(self, from_=0, to=100, orient=tk.HORIZONTAL,
                                      label="Master Volume", command=self.on_volume_change)
        self.volume_slider.pack(pady=5)

        # Initialize slider position and status
        self.update_volume_slider()
        self.update_status()

    def on_device_selected(self, event):
        selection = self.device_combo.current()
        device_ptr, name = self.devices[selection]
        # Activate the IAudioEndpointVolume interface for the selected device
        self.audio_volume = activate_audio_endpoint_volume(device_ptr)
        self.update_volume_slider()
        self.update_status()

    def set_as_default(self):
        selection = self.device_combo.current()
        device_ptr, name = self.devices[selection]
        switch_default_device(device_ptr)

    def volume_up(self):
        volume_step_up(self.audio_volume)
        self.update_volume_slider()
        self.update_status()

    def volume_down(self):
        volume_step_down(self.audio_volume)
        self.update_volume_slider()
        self.update_status()

    def toggle_mute(self):
        current = get_mute(self.audio_volume)
        set_mute(self.audio_volume, not current)
        self.update_status()

    def on_volume_change(self, value):
        # Slider callback: set master volume (value is a string, convert to float)
        vol = float(value) / 100.0
        set_master_volume(self.audio_volume, vol)
        self.update_status()

    def update_volume_slider(self):
        # Update the slider to reflect current master volume (0.0 to 1.0 converted to 0-100)
        current_volume = get_master_volume(self.audio_volume)
        self.volume_slider.set(int(current_volume * 100))

    def update_status(self):
        muted = get_mute(self.audio_volume)
        current_volume = get_master_volume(self.audio_volume)
        status = f"Status: {'Muted' if muted else 'Unmuted'}, Volume: {int(current_volume * 100)}%"
        self.lbl_status.config(text=status)

# ============================================================
# Main Function: Execute Steps and Launch GUI
# ============================================================
def main():
    # Step 1: Initialize COM
    init_com()
    
    # Step 2: Create the IMMDeviceEnumerator object
    enumerator = create_device_enumerator()
    
    # Step 3: Get the default audio endpoint
    default_endpoint = get_default_endpoint(enumerator)
    
    # Launch the Tkinter GUI with device selection, volume control, slider, and default switch
    app = VolumeControlApp(enumerator, default_endpoint)
    app.mainloop()

if __name__ == "__main__":
    main()
