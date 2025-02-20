import sys
import subprocess
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL

# Auto-install required packages
def install_dependencies():
    try:
        import pycaw
        import comtypes
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pycaw", "comtypes"])
        import pycaw
        import comtypes

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

def change_volume(volume: float):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume_control = cast(interface, POINTER(IAudioEndpointVolume))
    
    if volume == 0:
        volume_control.SetMute(1, None)
    else:
        volume_control.SetMute(0, None)
        volume_control.SetMasterVolumeLevelScalar(volume, None)

def main():
    install_dependencies()
    
    if len(sys.argv) != 2:
        print("Usage: python change_volume.py <volume_percentage>")
        return
    
    try:
        volume_percent = int(sys.argv[1])
        if 0 <= volume_percent <= 100:
            change_volume(volume_percent / 100)
        else:
            print("Error: Volume percentage must be between 0 and 100.")
    except ValueError:
        print("Error: Invalid number format.")

if __name__ == "__main__":
    main()
