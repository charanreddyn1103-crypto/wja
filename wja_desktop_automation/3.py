# base_class.py
import json
import logging
import time
from pywinauto import Application, Desktop
from pywinauto.timings import TimeoutError

class BaseClass:
    def __init__(self):
        with open("config.json", "r") as f:
            self.config = json.load(f)
        self.app = None
        self.desktop = Desktop(backend=self.config["backend"])
        self.main_window = None
        self.logger = self._get_logger()

    def _get_logger(self):
        logger = logging.getLogger("WJA_Automation")
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                "%(asctime)s %(levelname)s %(message)s"
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        return logger

    def launch_application(self):
        try:
            self.logger.info("Launching HP Web Jetadmin...")
            self.app = Application(backend=self.config["backend"]).start(self.config["application_path"])
            self.main_window = self.connect_to_main_window()
            self.logger.info("Application launched and main window connected.")
        except Exception as e:
            self.logger.error(f"Failed to launch application: {e}")
            raise

    def connect_to_main_window(self):
        try:
            self.logger.info("Connecting to main window...")
            window = self.desktop.window(title_re=".*Web Jetadmin.*")
            window.wait("visible", timeout=self.config["default_timeout"])
            window.set_focus()
            return window
        except TimeoutError:
            self.logger.error("Main window not found within timeout.")
            raise
        except Exception as e:
            self.logger.error(f"Error connecting to main window: {e}")
            raise

    def wait_for_element(self, parent, **locator):
        timeout = self.config.get("default_timeout", 20)
        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                element = parent.child_window(**locator)
                element.wait("exists ready visible", timeout=1)
                return element
            except Exception:
                time.sleep(0.5)
        self.logger.error(f"Element with locator {locator} not found within timeout.")
        raise TimeoutError(f"Element with locator {locator} not found.")

    def click_element(self, parent, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            element.click_input()
            self.logger.info(f"Clicked element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to click element {locator}: {e}")
            raise

    def enter_text(self, parent, text, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            element.type_keys("^a{BACKSPACE}", set_foreground=True)
            element.type_keys(text, with_spaces=True, set_foreground=True)
            self.logger.info(f"Entered text '{text}' in element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to enter text in element {locator}: {e}")
            raise

    def set_checkbox(self, parent, check, **locator):
        try:
            element = self.wait_for_element(parent, **locator)
            element.set_focus()
            current = element.get_toggle_state()
            if (check and current == 0) or (not check and current == 1):
                element.toggle()
                self.logger.info(f"Checkbox {locator} set to {'checked' if check else 'unchecked'}.")
            else:
                self.logger.info(f"Checkbox {locator} already in desired state.")
        except Exception as e:
            self.logger.error(f"Failed to set checkbox {locator}: {e}")
            raise

    def select_combobox(self, parent, value, **locator):
        try:
            combobox = self.wait_for_element(parent, **locator)
            combobox.set_focus()
            combobox.select(value)
            self.logger.info(f"Selected '{value}' in combobox: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to select '{value}' in combobox {locator}: {e}")
            raise

# test.py
import pytest
import allure
from base_class import BaseClass

@allure.feature("Edit OXPd Statistics Agent Record")
class TestEditOXPdStatsAgent(BaseClass):

    @pytest.fixture(autouse=True, scope="class")
    def setup_class(self):
        self.launch_application()
        yield
        if self.app:
            self.app.kill()

    @allure.story("Edit OXPd Statistics Agent Record")
    def test_edit_oxpd_statistics_agent(self):
        """
        Test Case: Edit OXPd Statistics Agent Record
        Steps:
        1. Navigate to OXPd Statistics Agents screen.
        2. Select agent record and click Edit.
        3. Edit fields and save.
        """
        # Step 1: Navigate to 'Tools > Options > Device Management > Configuration > OXPd Statistics Agents'
        with allure.step("Navigate to OXPd Statistics Agents screen"):
            self._navigate_to_oxpd_statistics_agents()

        # Step 2: Select agent record and click Edit
        with allure.step("Select agent record and click Edit"):
            agents_window = self.main_window.child_window(title="OXPd Statistics Agents", control_type="Window")
            # Select first row in the table (assuming DataGrid)
            self.click_element(agents_window, control_type="DataItem", found_index=0)
            self.click_element(agents_window, title="Edit", control_type="Button")

        # Step 3: Edit fields in the Edit dialog
        edit_dialog = self.desktop.window(title_re=".*Edit OXPd Statistics Agent.*")
        edit_dialog.wait("visible", timeout=self.config["default_timeout"])
        edit_dialog.set_focus()

        # Enter Name
        with allure.step("Enter agent name (max 100 chars)"):
            agent_name = "AutomationAgent001"
            self.enter_text(edit_dialog, agent_name, title="Name", control_type="Edit")

        # Set/Clear Critical agent checkbox
        with allure.step("Set/Clear 'Critical agent' checkbox"):
            self.set_checkbox(edit_dialog, check=True, title="Critical agent (acknowledgment required for delete)", control_type="CheckBox")

        # Select Data persistence frequency
        with allure.step("Select Data persistence frequency"):
            self.select_combobox(edit_dialog, "Job", title="Data persistence frequency", control_type="ComboBox")

        # Enter URI
        with allure.step("Enter agent URI (max 256 chars, valid protocol)"):
            uri = "https://oxpd-agent-server.example.com"
            self.enter_text(edit_dialog, uri, title="URI", control_type="Edit")

        # Save changes
        with allure.step("Click OK to save changes"):
            self.click_element(edit_dialog, title="OK", control_type="Button")

        # Verify dialog closed and record updated
        with allure.step("Verify record updated and dialog closed"):
            assert not edit_dialog.exists(timeout=5), "Edit dialog did not close after saving."
            # Optionally, verify the updated values in the agents list

    def _navigate_to_oxpd_statistics_agents(self):
        """Navigate the menu: Tools > Options > Device Management > Configuration > OXPd Statistics Agents"""
        # Menu navigation may use menu bar or tree view, adjust as per actual UI
        try:
            self.logger.info("Navigating to OXPd Statistics Agents via menu...")
            menu_bar = self.main_window.child_window(control_type="MenuBar")
            self.click_element(menu_bar, title="Tools", control_type="MenuItem")
            tools_menu = self.main_window.child_window(title="Tools", control_type="Menu")
            self.click_element(tools_menu, title="Options", control_type="MenuItem")
            options_menu = self.main_window.child_window(title="Options", control_type="Menu")
            self.click_element(options_menu, title="Device Management", control_type="MenuItem")
            device_mgmt_menu = self.main_window.child_window(title="Device Management", control_type="Menu")
            self.click_element(device_mgmt_menu, title="Configuration", control_type="MenuItem")
            config_menu = self.main_window.child_window(title="Configuration", control_type="Menu")
            self.click_element(config_menu, title="OXPd Statistics Agents", control_type="MenuItem")
            # Wait for the OXPd Statistics Agents window to appear
            agents_window = self.desktop.window(title="OXPd Statistics Agents")
            agents_window.wait("visible", timeout=self.config["default_timeout"])
            agents_window.set_focus()
            self.logger.info("OXPd Statistics Agents window is visible.")
        except Exception as e:
            self.logger.error(f"Failed to navigate to OXPd Statistics Agents: {e}")
            raise

# config.json
{
  "application_path": "C:/Program Files/HP/Web Jetadmin/wja.exe",
  "backend": "uia",
  "default_timeout": 20,
  "retry_count": 3
}

# requirements.txt
pywinauto
pytest
pytest-html
allure-pytest
