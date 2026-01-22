# base_class.py
import json
import logging
import time
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

class BaseClass:
    def __init__(self):
        with open("config.json") as f:
            self.config = json.load(f)
        self.app = Application(backend=self.config["backend"])
        self.desktop = Desktop(backend=self.config["backend"])
        self.main_window = None
        self.logger = self._get_logger()

    def _get_logger(self):
        logger = logging.getLogger("WJA_Automation")
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        return logger

    def launch_application(self):
        try:
            self.logger.info("Launching HP Web Jetadmin application...")
            self.app.start(self.config["application_path"])
            self.main_window = self.desktop.window(title_re=".*Web Jetadmin.*")
            self.main_window.wait("visible", timeout=self.config["default_timeout"])
            self.main_window.set_focus()
            self.logger.info("Application launched and main window focused.")
        except Exception as e:
            self.logger.error(f"Failed to launch application: {e}")
            raise

    def connect_to_main_window(self):
        try:
            self.main_window = self.desktop.window(title_re=".*Web Jetadmin.*")
            self.main_window.wait("visible", timeout=self.config["default_timeout"])
            self.main_window.set_focus()
            self.logger.info("Connected to main window.")
        except Exception as e:
            self.logger.error(f"Failed to connect to main window: {e}")
            raise

    def wait_for_element(self, parent, **locator):
        try:
            self.logger.info(f"Waiting for element: {locator}")
            element = parent.child_window(**locator)
            element.wait("exists enabled visible ready", timeout=self.config["default_timeout"])
            self.logger.info(f"Element found: {locator}")
            return element
        except Exception as e:
            self.logger.error(f"Element not found: {locator} - {e}")
            raise

    def click_element(self, parent, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            element.click_input()
            self.logger.info(f"Clicked element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to click element: {locator} - {e}")
            raise

    def enter_text(self, parent, text, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            element.type_keys(text, with_spaces=True)
            self.logger.info(f"Entered text '{text}' in element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to enter text in element: {locator} - {e}")
            raise

    def right_click_element(self, parent, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            element.right_click_input()
            self.logger.info(f"Right-clicked element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to right-click element: {locator} - {e}")
            raise

    def select_context_menu(self, menu_title):
        try:
            # Assumes context menu is open and visible
            menu = self.desktop.window(control_type="Menu")
            item = self.wait_for_element(menu, title=menu_title, control_type="MenuItem")
            item.click_input()
            self.logger.info(f"Selected context menu item: {menu_title}")
        except Exception as e:
            self.logger.error(f"Failed to select context menu item '{menu_title}': {e}")
            raise

    def wizard_next(self, wizard_window):
        try:
            next_btn = self.wait_for_element(wizard_window, title="Next", control_type="Button")
            next_btn.click_input()
            self.logger.info("Clicked 'Next' in wizard.")
        except Exception as e:
            self.logger.error(f"Failed to click 'Next' in wizard: {e}")
            raise

    def wizard_finish(self, wizard_window):
        try:
            finish_btn = self.wait_for_element(wizard_window, title="Finish", control_type="Button")
            finish_btn.click_input()
            self.logger.info("Clicked 'Finish' in wizard.")
        except Exception as e:
            self.logger.error(f"Failed to click 'Finish' in wizard: {e}")
            raise

    def bring_window_to_front(self, window):
        try:
            window.set_focus()
            self.logger.info("Window brought to front.")
        except Exception as e:
            self.logger.warning(f"Failed to bring window to front: {e}")
