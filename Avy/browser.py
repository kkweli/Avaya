import os

from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


def driverOpt(browser):
    if browser == 'ie':
        cap = DesiredCapabilities.INTERNETEXPLORER.copy()
        # cap['platform'] = "Win8"
        cap['version'] = "11"
        cap['browserName'] = "internet explorer"
        cap['ignoreProtectedModeSettings'] = True
        cap['nativeEvents'] = False
        cap['requireWindowFocus'] = True
        cap['INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS'] = True
        ie_options = webdriver.IeOptions()
        ie_options.add_argument('-private')
        # ie_options.force_create_process_api = True #ie 8 or higher only
        ie_options.ignore_protected_mode_settings = True
        ie_options.ignore_zoom_level = True
        ie_options.add_argument("--ignore-certificate-errors")
        ie_options.add_argument("--disable-infobars")
        ie_options.add_argument("--start-minimized")
        ie_options.add_argument("--disable-extensions")
        ie_options.binary_location = os.path.join('C:', 'Program Files', 'internet explorer', 'iexplore.exe')
        return webdriver.Ie(capabilities=cap,
                            executable_path=os.path.abspath("IEDriverServer"),
                            options=ie_options)


class Browser:
    pass
