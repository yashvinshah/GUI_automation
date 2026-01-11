"""
macOS-specific GUI automation functions.
"""

import subprocess
import logging
import platform
from typing import Optional

logger = logging.getLogger(__name__)


class MacGUI:
    """macOS GUI automation implementation."""
    
    def __init__(self):
        if platform.system() != "Darwin":
            raise RuntimeError("MacGUI can only be used on macOS")
    
    def open_application_with_file(self, app_name: str, file_path: str) -> bool:
        """
        Open an application with a specific file on macOS.
        
        Args:
            app_name: Name of the application (e.g., "Microsoft Excel")
            file_path: Path to the file to open
            
        Returns:
            True if successful
        """
        try:
            subprocess.Popen(["open", "-a", app_name, file_path])
            return True
        except Exception as e:
            logger.error(f"Error opening {app_name} with file {file_path}: {e}")
            return False
    
    def open_url_in_browser(self, browser_name: str, url: str) -> bool:
        """
        Open a URL in a specific browser on macOS.
        
        Args:
            browser_name: Name of the browser (e.g., "Google Chrome")
            url: URL to open
            
        Returns:
            True if successful
        """
        try:
            subprocess.Popen(["open", "-a", browser_name, url])
            return True
        except Exception as e:
            logger.error(f"Error opening {url} in {browser_name}: {e}")
            return False
    
    def send_keystroke_to_app(self, app_name: str, key: str, modifiers: Optional[list] = None) -> bool:
        """
        Send a keystroke to a specific application using AppleScript.
        
        Args:
            app_name: Name of the application
            key: Key to press (e.g., "f")
            modifiers: List of modifier keys (e.g., ["command"])
            
        Returns:
            True if successful
        """
        try:
            modifier_str = ""
            if modifiers:
                if "command" in modifiers:
                    modifier_str = 'using command down'
                elif "control" in modifiers:
                    modifier_str = 'using control down'
                elif "option" in modifiers:
                    modifier_str = 'using option down'
            
            if modifier_str:
                script = f'''tell application "{app_name}"
    activate
end tell
delay 0.2
tell application "System Events"
    tell process "{app_name}"
        keystroke "{key}" {modifier_str}
    end tell
end tell'''
            else:
                script = f'''tell application "{app_name}"
    activate
end tell
delay 0.2
tell application "System Events"
    tell process "{app_name}"
        keystroke "{key}"
    end tell
end tell'''
            subprocess.run(["osascript", "-e", script], check=True, timeout=3)
            return True
        except Exception as e:
            logger.warning(f"AppleScript keystroke failed: {e}")
            return False
    
    def check_app_active(self, app_name: str) -> bool:
        """
        Check if an application is currently active/frontmost.
        
        Args:
            app_name: Name of the application
            
        Returns:
            True if application is active
        """
        try:
            script = f'''
            tell application "System Events"
                set frontApp to name of first application process whose frontmost is true
                return frontApp contains "{app_name}"
            end tell
            '''
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=3
            )
            return "true" in result.stdout.lower()
        except Exception as e:
            logger.warning(f"Error checking app active: {e}")
            return False
