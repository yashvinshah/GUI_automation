"""
Window Manager for GUI Automation
Handles window focusing, error detection, retries, and logging.
"""

import pyautogui
import time
import logging
from pathlib import Path
from typing import Optional, Tuple, Callable
import subprocess
import platform

# Configure logging - use project root directory
PROJECT_ROOT = Path(__file__).parent.parent
LOG_DIR = PROJECT_ROOT / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / "window_manager.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Screenshot directory for debugging - use same logs directory
SCREENSHOT_DIR = LOG_DIR
SCREENSHOT_DIR.mkdir(exist_ok=True)


class WindowManager:
    """
    Manages window operations with robust error handling, retries, and logging.
    """
    
    def __init__(self, app_name: str, max_retries: int = 3, wait_timeout: float = 10.0):
        """
        Initialize WindowManager.
        
        Args:
            app_name: Name of the application (e.g., "Microsoft Excel", "Google Chrome")
            max_retries: Maximum number of retries for operations
            wait_timeout: Maximum time to wait for conditions (seconds)
        """
        self.app_name = app_name
        self.max_retries = max_retries
        self.wait_timeout = wait_timeout
        self.screen_width, self.screen_height = pyautogui.size()
        logger.info(f"WindowManager initialized for {app_name} on {self.screen_width}x{self.screen_height} screen")
    
    def take_screenshot(self, filename: str) -> Path:
        """
        Take a screenshot for debugging purposes.
        
        Args:
            filename: Name of the screenshot file
            
        Returns:
            Path to the screenshot file
        """
        screenshot_path = SCREENSHOT_DIR / f"{filename}_{int(time.time())}.png"
        try:
            pyautogui.screenshot(str(screenshot_path))
            logger.debug(f"Screenshot saved: {screenshot_path}")
            return screenshot_path
        except Exception as e:
            logger.error(f"Failed to take screenshot: {e}")
            return screenshot_path
    
    def wait_for_condition(
        self,
        condition_func: Callable[[], bool],
        condition_name: str,
        timeout: Optional[float] = None,
        check_interval: float = 0.5
    ) -> bool:
        """
        Wait for a condition to become true with timeout.
        
        Args:
            condition_func: Function that returns True when condition is met
            condition_name: Description of the condition (for logging)
            timeout: Maximum time to wait (uses self.wait_timeout if None)
            check_interval: Time between condition checks
            
        Returns:
            True if condition met, False if timeout
        """
        timeout = timeout or self.wait_timeout
        start_time = time.time()
        
        logger.info(f"Waiting for condition: {condition_name}")
        
        while time.time() - start_time < timeout:
            try:
                if condition_func():
                    elapsed = time.time() - start_time
                    logger.info(f"Condition met: {condition_name} (took {elapsed:.2f}s)")
                    return True
            except Exception as e:
                logger.warning(f"Error checking condition {condition_name}: {e}")
            
            time.sleep(check_interval)
        
        elapsed = time.time() - start_time
        logger.warning(f"Timeout waiting for condition: {condition_name} (waited {elapsed:.2f}s)")
        self.take_screenshot(f"timeout_{condition_name}")
        return False
    
    def focus_application(self) -> bool:
        """
        Focus on the target application window.
        
        Returns:
            True if successful, False otherwise
        """
        logger.info(f"Attempting to focus on {self.app_name}")
        
        if platform.system() == "Darwin":  # macOS
            try:
                # Use AppleScript to activate the application
                script = f'''
                tell application "{self.app_name}"
                    activate
                end tell
                '''
                subprocess.run(["osascript", "-e", script], check=True, timeout=5)
                
                # Wait for window to be active
                def is_app_active():
                    script_check = f'''
                    tell application "System Events"
                        set frontApp to name of first application process whose frontmost is true
                        return frontApp contains "{self.app_name}"
                    end tell
                    '''
                    result = subprocess.run(
                        ["osascript", "-e", script_check],
                        capture_output=True,
                        text=True,
                        timeout=3
                    )
                    return "true" in result.stdout.lower()
                
                if self.wait_for_condition(is_app_active, f"{self.app_name} window active", timeout=5):
                    logger.info(f"Successfully focused on {self.app_name}")
                    time.sleep(0.5)  # Small delay for window to fully activate
                    return True
                else:
                    logger.error(f"Failed to verify {self.app_name} is active")
                    self.take_screenshot("focus_failed")
                    return False
                    
            except subprocess.TimeoutExpired:
                logger.error(f"Timeout activating {self.app_name}")
                self.take_screenshot("focus_timeout")
                return False
            except Exception as e:
                logger.error(f"Error focusing {self.app_name}: {e}")
                self.take_screenshot("focus_error")
                return False
        else:
            logger.warning(f"Platform {platform.system()} not fully supported for window focusing")
            return True
    
    def click_with_retry(
        self,
        x: int,
        y: int,
        clicks: int = 1,
        button: str = "left",
        verify_func: Optional[Callable[[], bool]] = None
    ) -> bool:
        """
        Click at coordinates with retry logic.
        
        Args:
            x, y: Click coordinates
            clicks: Number of clicks
            button: Mouse button ("left" or "right")
            verify_func: Optional function to verify click success
            
        Returns:
            True if successful, False otherwise
        """
        for attempt in range(1, self.max_retries + 1):
            try:
                logger.debug(f"Click attempt {attempt}/{self.max_retries} at ({x}, {y})")
                pyautogui.click(x, y, clicks=clicks, button=button)
                time.sleep(0.3)
                
                if verify_func:
                    if verify_func():
                        logger.info(f"Click verified at ({x}, {y})")
                        return True
                    else:
                        logger.warning(f"Click verification failed at ({x}, {y}), attempt {attempt}")
                else:
                    return True
                    
            except Exception as e:
                logger.warning(f"Click error at ({x}, {y}), attempt {attempt}: {e}")
            
            if attempt < self.max_retries:
                time.sleep(1)
                self.take_screenshot(f"click_retry_{attempt}")
        
        logger.error(f"Failed to click at ({x}, {y}) after {self.max_retries} attempts")
        self.take_screenshot("click_failed")
        return False
    
    def click_center_safe(self) -> bool:
        """
        Click in the center of the screen (safe area) to ensure focus.
        Uses a safe area that avoids screen edges and common UI elements.
        
        Returns:
            True if successful
        """
        # Use center area, avoiding edges
        safe_x = int(self.screen_width * 0.5)
        safe_y = int(self.screen_height * 0.5)
        
        logger.info(f"Clicking safe center area at ({safe_x}, {safe_y})")
        return self.click_with_retry(safe_x, safe_y)
    
    def execute_with_retry(
        self,
        operation: Callable[[], bool],
        operation_name: str,
        verify_func: Optional[Callable[[], bool]] = None
    ) -> bool:
        """
        Execute an operation with retry logic.
        
        Args:
            operation: Function to execute (returns True on success)
            operation_name: Description of operation (for logging)
            verify_func: Optional function to verify operation success
            
        Returns:
            True if successful, False otherwise
        """
        logger.info(f"Executing operation: {operation_name}")
        
        for attempt in range(1, self.max_retries + 1):
            try:
                if operation():
                    if verify_func:
                        if verify_func():
                            logger.info(f"Operation successful: {operation_name}")
                            return True
                        else:
                            logger.warning(f"Operation verification failed: {operation_name}, attempt {attempt}")
                    else:
                        logger.info(f"Operation successful: {operation_name}")
                        return True
                else:
                    logger.warning(f"Operation returned False: {operation_name}, attempt {attempt}")
                    
            except Exception as e:
                logger.error(f"Operation error: {operation_name}, attempt {attempt}: {e}")
                self.take_screenshot(f"operation_error_{operation_name}_{attempt}")
            
            if attempt < self.max_retries:
                wait_time = attempt * 1.0  # Exponential backoff
                logger.info(f"Retrying {operation_name} after {wait_time}s...")
                time.sleep(wait_time)
        
        logger.error(f"Operation failed after {self.max_retries} attempts: {operation_name}")
        self.take_screenshot(f"operation_failed_{operation_name}")
        return False
    
    def wait_for_window_ready(self, timeout: Optional[float] = None) -> bool:
        """
        Wait for the application window to be ready for interaction.
        
        Args:
            timeout: Maximum time to wait
            
        Returns:
            True if window is ready
        """
        def window_ready():
            # Focus the application first
            if not self.focus_application():
                return False
            
            # Click center to ensure focus
            self.click_center_safe()
            time.sleep(0.5)
            
            # Check if application is still active
            if platform.system() == "Darwin":
                try:
                    script = '''
                    tell application "System Events"
                        set frontApp to name of first application process whose frontmost is true
                        return frontApp
                    end tell
                    '''
                    result = subprocess.run(
                        ["osascript", "-e", script],
                        capture_output=True,
                        text=True,
                        timeout=3
                    )
                    app_active = self.app_name.lower() in result.stdout.lower()
                    return app_active
                except:
                    return False
            return True
        
        return self.wait_for_condition(
            window_ready,
            f"{self.app_name} window ready",
            timeout=timeout or self.wait_timeout
        )
    
    def log_operation(self, message: str, level: str = "INFO"):
        """
        Log an operation with appropriate level.
        
        Args:
            message: Log message
            level: Log level (INFO, WARNING, ERROR, DEBUG)
        """
        log_func = getattr(logger, level.lower(), logger.info)
        log_func(message)
