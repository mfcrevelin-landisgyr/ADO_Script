import platformdirs as dirs
import pickle as pk
import subprocess
import shutil
import time
import sys
import os

try:
	import selenium
except ImportError:
	subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium"])
	import selenium

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException,TimeoutException, NoSuchElementException


CACHE_FILE_PATH = os.path.join(dirs.user_cache_dir("ADO_Script"), "cache.pk")

cache = {}
variables = {}
driver = None

def load_cache():
	global cache
	if os.path.isfile(CACHE_FILE_PATH):
		with open(CACHE_FILE_PATH, "rb") as file:
			cache = pk.load(file)

def save_cache():
	global cache
	os.makedirs(os.path.dirname(CACHE_FILE_PATH), exist_ok=True)
	with open(CACHE_FILE_PATH, "wb") as file:
		pk.dump(cache, file)

def get_browser_path():
	global cache

	browsers = {
		"Firefox": "firefox.exe",
		"Chrome": "chrome.exe",
		"Edge": "msedge.exe"
	}

	search_paths = [path for path in [
		os.environ.get("ProgramFiles"),
		os.environ.get("ProgramFiles(x86)"),
		os.environ.get("LocalAppData")
	] if not path == None]

	browser_paths = {}

	for browser, executable in browsers.items():
		for base_path in search_paths:
			for root, dirs, files in os.walk(base_path):
				if executable in files:
					browser_paths[browser] = os.path.join(root, executable)
					break
			if browser in browser_paths:
				break

	if len(browser_paths) == 0:
		print(f"\033[91mNo supported browsers installed. (Firefox, Chrome, Edge)\033[0m")
		sys.exit(1)

	if len(browser_paths)>1:
		print("\033[33mMultiple browsers detected. Please select one:\033[0m")
		cached_browser = cache.get("browser",None)
		found_browsers = list(browser_paths.keys())
		for i, browser in enumerate(found_browsers, start=1):
			default_tag = " *" if browser == cached_browser else ""
			print(f"  {i}) {browser}{default_tag}")

		while True:
			choice = input(": ").strip()
			if choice=="" and cached_browser:
				break

			if choice.isdigit() and 1 <= int(choice) <= len(found_browsers):
				cache["browser"] = found_browsers[int(choice) - 1]
				cache["browser_path"] = browser_paths[cache["browser"]]
				break
			else:
				sys.stdout.write("\x1b[1A\x1b[2K")
				sys.stdout.flush()
		
		save_cache()

def start_driver():
	global driver
	global cache

	browser_name = cache["browser"] 
	browser_binary_path = cache["browser_path"]
	browser_driver = {
		"Firefox": "geckodriver.exe",
		"Chrome": "chromedriver.exe",
		"Edge": "msedgedriver.exe"
	}[browser_name]

	if not os.path.isfile(browser_driver):
		print( f"\033[91m\nError: {browser_driver} not found.\033[33m\nPlease download the driver and place it in the same folder as this script.\033[0m")
		sys.exit(1)

	if browser_name == "Firefox":
		service = webdriver.FirefoxService(executable_path=browser_driver)
		options = webdriver.FirefoxOptions()
	elif browser_name == "Chrome":
		service = webdriver.ChromeService(executable_path=browser_driver)
		options = webdriver.ChromeOptions()
	elif browser_name == "Edge":
		service = webdriver.EdgeService(executable_path=browser_driver)
		options = webdriver.EdgeOptions()
	else:
		print( f"\033[91m\nUnexpected driver name.\033[0m")
		sys.exit(1)

	options.binary_location = browser_binary_path
	driver = webdriver.__dict__[browser_name](service=service, options=options)


def main():
	global driver
	global main_tab
	global cache

	load_cache()
	get_browser_path()

	start_driver()

	driver.get("https://global.tfs.landisgyr.net/tfs/DefaultCollection_Brazil/HQA/_testPlans")
	input("\033[33mPress ENTER to continue:\033[0m")

	driver.quit()

if __name__ == "__main__":
	main()