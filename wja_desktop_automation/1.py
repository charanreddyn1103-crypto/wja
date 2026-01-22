# base_class.py
import json
import logging
import time
from pywinauto import Application, Desktop
from pywinauto.timings import TimeoutError
import allure

class BaseClass:
    def __init__(self):
        with open("config.json", "r") as f:
            self.config = json.load(f)
        self.app = None
        self.main_window = None
        self.desktop = Desktop(backend=self.config["backend"])
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

    @allure.step("Launching HP Web Jetadmin application")
    def launch_application(self):
        try:
            self.logger.info("Launching application: %s", self.config["application_path"])
            self.app = Application(backend=self.config["backend"]).start(self.config["application_path"])
            self.main_window = self.desktop.window(title_re=".*Web Jetadmin.*")
            self.main_window.wait("visible", timeout=self.config["default_timeout"])
            self.main_window.set_focus()
            self.logger.info("Application launched and main window focused.")
        except Exception as e:
            self.logger.error(f"Failed to launch application: {e}")
            allure.attach(str(e), name="Launch Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Connecting to main window")
    def connect_to_main_window(self):
        try:
            self.main_window = self.desktop.window(title_re=".*Web Jetadmin.*")
            self.main_window.wait("visible", timeout=self.config["default_timeout"])
            self.main_window.set_focus()
            self.logger.info("Connected to main window.")
        except Exception as e:
            self.logger.error(f"Failed to connect to main window: {e}")
            allure.attach(str(e), name="Connect Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Waiting for element: {locator}")
    def wait_for_element(self, locator, timeout=None):
        timeout = timeout or self.config["default_timeout"]
        try:
            ctrl = self.main_window.child_window(**locator)
            ctrl.wait("exists enabled visible ready", timeout=timeout)
            self.logger.info(f"Element found: {locator}")
            return ctrl
        except TimeoutError:
            self.logger.error(f"Timeout waiting for element: {locator}")
            allure.attach(str(locator), name="Wait Timeout", attachment_type=allure.attachment_type.TEXT)
            raise
        except Exception as e:
            self.logger.error(f"Error waiting for element {locator}: {e}")
            allure.attach(str(e), name="Wait Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Clicking element: {locator}")
    def click_element(self, locator):
        try:
            ctrl = self.wait_for_element(locator)
            ctrl.set_focus()
            ctrl.click_input()
            self.logger.info(f"Clicked element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to click element {locator}: {e}")
            allure.attach(str(e), name="Click Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Entering text '{text}' into element: {locator}")
    def enter_text(self, locator, text, clear=True):
        try:
            ctrl = self.wait_for_element(locator)
            ctrl.set_focus()
            if clear:
                ctrl.type_keys('^a{BACKSPACE}')
            ctrl.type_keys(text, with_spaces=True, set_foreground=True)
            self.logger.info(f"Entered text '{text}' into element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to enter text into element {locator}: {e}")
            allure.attach(str(e), name="Enter Text Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Expanding tree item: {locator}")
    def expand_tree_item(self, locator):
        try:
            ctrl = self.wait_for_element(locator)
            if not ctrl.is_expanded():
                ctrl.expand()
            self.logger.info(f"Expanded tree item: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to expand tree item {locator}: {e}")
            allure.attach(str(e), name="Expand Error", attachment_type=allure.attachment_type.TEXT)
            raise

    @allure.step("Right-clicking element: {locator}")
    def right_click_element(self, locator):
        try:
            ctrl = self.wait_for_element(locator)
            ctrl.set_focus()
            ctrl.right_click_input()
            self.logger.info(f"Right-clicked element: {locator}")
        except Exception as e:
            self.logger.error(f"Failed to right-click element {locator}: {e}")
            allure.attach(str(e), name="Right Click Error", attachment_type=allure.attachment_type.TEXT)
            raise

# test.py
import pytest
import allure
import time
from base_class import BaseClass

@pytest.mark.desktop
@allure.feature("Automatic Group Creation")
@allure.story("Verification of Automatic Group Creation")
def test_automatic_group_creation():
    """
    Test Case: Verification of Automatic Group Creation in HP Web Jetadmin
    Steps:
    1. Launch WJA and open the Groups tree.
    2. Right-click 'Groups' and select 'New group'.
    3. Enter a unique group name, select parent group, and set as Automatic group.
    4. Define filter criteria (Basic).
    5. Complete the wizard and verify group creation and device listing.
    """
    base = BaseClass()
    base.launch_application()

    # Step 1: Expand Groups tree
    groups_tree = {"title": "Groups", "control_type": "TreeItem"}
    base.expand_tree_item(groups_tree)

    # Step 2: Right-click 'Groups' and select 'New group'
    base.right_click_element(groups_tree)
    new_group_menu = {"title": "New group", "control_type": "MenuItem"}
    base.click_element(new_group_menu)

    # Step 3: Specify group options
    wizard_window = base.wait_for_element({"title_re": "Specify group options.*", "control_type": "Window"})
    group_name_edit = {"auto_id": "txtGroupName", "control_type": "Edit"}
    unique_group_name = f"AutoGroup_{int(time.time())}"
    base.enter_text(group_name_edit, unique_group_name)

    automatic_group_checkbox = {"title": "Automatic group", "control_type": "CheckBox"}
    base.click_element(automatic_group_checkbox)

    next_button = {"title": "Next >", "control_type": "Button"}
    base.click_element(next_button)

    # Step 4: Specify filter criteria (Basic)
    filter_criteria_window = base.wait_for_element({"title_re": "Specify filter criteria.*", "control_type": "Window"})
    basic_radio = {"title": "Basic", "control_type": "RadioButton"}
    base.click_element(basic_radio)

    add_button = {"title": "Add", "control_type": "Button"}
    base.click_element(add_button)

    # Function page
    device_property_combo = {"auto_id": "cmbDeviceProperty", "control_type": "ComboBox"}
    base.click_element(device_property_combo)
    # Select first property (for demo, just send keys)
    base.main_window.child_window(**device_property_combo).type_keys("{DOWN}{ENTER}")

    filter_function_combo = {"auto_id": "cmbFilterFunction", "control_type": "ComboBox"}
    base.click_element(filter_function_combo)
    base.main_window.child_window(**filter_function_combo).type_keys("{DOWN}{ENTER}")

    value_edit = {"auto_id": "txtValue", "control_type": "Edit"}
    base.enter_text(value_edit, "HP")

    ok_button = {"title": "OK", "control_type": "Button"}
    base.click_element(ok_button)

    # Save filter and proceed
    base.click_element(next_button)

    # Step 5: Specify group properties (skip, just next)
    group_properties_window = base.wait_for_element({"title_re": "Specify group properties.*", "control_type": "Window"})
    base.click_element(next_button)

    # Step 6: Configure group policies (skip, just finish)
    group_policies_window = base.wait_for_element({"title_re": "Configure group policies.*", "control_type": "Window"})
    finish_button = {"title": "Finish", "control_type": "Button"}
    base.click_element(finish_button)

    # Step 7: Verify group creation
    base.connect_to_main_window()
    base.expand_tree_item(groups_tree)
    created_group = {"title": unique_group_name, "control_type": "TreeItem"}
    group_ctrl = base.wait_for_element(created_group)
    assert group_ctrl.exists(), "Automatic group not found in Groups tree"

    # Step 8: Click on the created group and verify device details
    base.click_element(created_group)
    device_list = {"auto_id": "lstDevices", "control_type": "List"}
    device_list_ctrl = base.wait_for_element(device_list)
    assert device_list_ctrl.exists(), "Device list not displayed for automatic group"

    allure.attach(f"Automatic group '{unique_group_name}' created and devices listed.", name="Group Creation", attachment_type=allure.attachment_type.TEXT)

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
