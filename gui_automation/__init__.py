"""
GUI Automation package for platform-specific automation code.
"""

import platform

if platform.system() == "Darwin":
    from .mac_gui import MacGUI
    GUI = MacGUI()
elif platform.system() == "Windows":
    from .windows_gui import WindowsGUI
    GUI = WindowsGUI()
else:
    raise NotImplementedError(f"Platform {platform.system()} not supported")
