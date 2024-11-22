from config import *

from datetime import datetime
import pickle as pk
import subprocess
import hashlib
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

driver = None
configs_to_files = None
processed_entries = None
processed_rows = None
original_tab = None

main_table_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div/div/div/div[2]/table"
loading_dots_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div/div/div/div[1]/div[1]/span[2]"
new_table_xpath = "/html/body/div[2]/div/div/div/div/div/div[3]/div/div/div/table"
close_button_xpath = "/html/body/div[2]/div/div/div/div/div/div[2]/button"

entries_cache_path = "./processed_entries_cache.bin"
rows_cache_path = "./processed_rows_cache.bin"

def populate_configs_to_files():
	global configs_to_files
	configs_to_files = {}
	for configuration in configurations_mask:
		if configs_to_files.get(configuration,None) is None:
			configs_to_files[configuration] = {}
		walk_path = os.path.join(configurations_dir,configuration)
		if not os.path.isdir(walk_path):
			raise ValueError(f"Invalid path in mask: {walk_path}")
		for dir_path, dir_names, file_names in os.walk(walk_path):
			for file_name in file_names:
				if file_name.endswith('_.lst'):
					base_name = file_name.replace('_.lst','')
					if configs_to_files[configuration].get(base_name,None) is None:
						configs_to_files[configuration][base_name] = {}
					abs_file_path = os.path.join(dir_path,file_name)
					if not configs_to_files[configuration][base_name].get("short",None) is None:
						existing_file = configs_to_files[configuration][base_name]["short"]
						if os.path.getmtime(abs_file_path) > os.path.getmtime(existing_file):
							configs_to_files[configuration][base_name]["short"] = abs_file_path
					else:
						configs_to_files[configuration][base_name]["short"] = abs_file_path
				elif file_name.endswith('.lst'):
					base_name = file_name.replace('.lst','')
					if configs_to_files[configuration].get(base_name,None) is None:
						configs_to_files[configuration][base_name] = {}
					abs_file_path = os.path.join(dir_path,file_name)
					if not configs_to_files[configuration][base_name].get("long",None) is None:
						existing_file = configs_to_files[configuration][base_name]["long"]
						if os.path.getmtime(abs_file_path) > os.path.getmtime(existing_file):
							configs_to_files[configuration][base_name]["long"] = abs_file_path
					else:
						configs_to_files[configuration][base_name]["long"] = abs_file_path
		for base_name in configs_to_files[configuration]:
			configs_to_files[configuration][base_name] = list(configs_to_files[configuration][base_name].values())

def fetch_cached_files():
	global processed_entries
	global processed_rows

	if os.path.isfile(entries_cache_path):
		with open(entries_cache_path, "rb") as f:
			processed_entries = pk.load(f)
	else:
		processed_entries = set()
		with open(entries_cache_path, "wb") as f:
			pk.dump(processed_entries, f)

	if os.path.isfile(rows_cache_path):
		with open(rows_cache_path, "rb") as f:
			processed_rows = pk.load(f)
	else:
		processed_rows = set()
		with open(rows_cache_path, "wb") as f:
			pk.dump(processed_rows, f)

def start_driver():
	global driver

	if not os.path.isfile(browser_binary_path):
		raise ValueError(f"Browser binary not found: {browser_binary_path}")

	driver_paths = {
		"firefox": "geckodriver.exe",
		"chrome": "chromedriver.exe",
		"edge": "msedgedriver.exe"
	}

	driver_name = None
	for name, path in driver_paths.items():
		if name in browser_binary_path.lower():
			driver_name = name
			driver_path = path
			break

	if not driver_name:
		raise ValueError(f"Unsupported browser: {browser_binary_path}")

	if not os.path.isfile(driver_path):
		print( "\033[33m" + f"\nError: {driver_path} not found in the same folder as the script.\nPlease download the driver and place it in the same folder as this script." + "\033[0m")
		sys.exit(1)

	if driver_name == "firefox":
		service = webdriver.FirefoxService(executable_path=driver_path)
		options = webdriver.FirefoxOptions()
	elif driver_name == "chrome":
		service = webdriver.ChromeService(executable_path=driver_path)
		options = webdriver.ChromeOptions()
	elif driver_name == "edge":
		service = webdriver.EdgeService(executable_path=driver_path)
		options = webdriver.EdgeOptions()
	else:
		raise ValueError("Unexpected driver name.")

	options.binary_location = browser_binary_path
	driver = webdriver.__dict__[driver_name.capitalize()](service=service, options=options)
	return driver

def assess_table_parameters():
	global driver

	print("\033[33m" + f"\nAssessing page navigation parameters." + "\033[0m",end="")


	main_table = driver.find_element(By.XPATH, main_table_xpath)
	rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")

	if len(rows)==0:
		print("\033[91m" + "\n\nCould not find any rows in Test Points list. Aborting." + "\033[0m")
		driver.quit()
		sys.exit(1)

	try:
		script_name = rows[0].find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
		ActionChains(driver).move_to_element(script_name).double_click().perform()
		WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, new_table_xpath)))
		close_button = driver.find_element(By.XPATH, close_button_xpath)
		ActionChains(driver).move_to_element(close_button).click().perform()
	except TimeoutException:
		print("\033[91m" + "\n\nTest Case Results table took too long to load. Aborting." + "\033[0m")
		driver.quit()
		sys.exit(1)

	left = 0
	right = len(rows) - 1
	num_lines = -1
	while left <= right:
		mid = (left + right) // 2
		script_name = rows[mid].find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
		try:
			ActionChains(driver).move_to_element(script_name).double_click().perform()
			WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.XPATH, new_table_xpath)))
			close_button = driver.find_element(By.XPATH, close_button_xpath)
			ActionChains(driver).move_to_element(close_button).click().perform()

			num_lines = mid
			left = mid + 1
		except:
			right = mid - 1
		if left == right:
			num_lines = left

	main_table = driver.find_element(By.XPATH, main_table_xpath)
	rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
	script_names = [row.find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]") for row in rows]

	starting_len_rows = len(rows)
	for mid_index, script_name in enumerate(script_names):
		ActionChains(driver).move_to_element(script_name).click().send_keys(Keys.PAGE_DOWN).perform()
		time.sleep(.01)
		main_table = driver.find_element(By.XPATH, main_table_xpath)
		rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
		if not len(rows) == starting_len_rows:
			break
	mid_index-=1

	main_table = driver.find_element(By.XPATH, main_table_xpath)
	rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")

	skip_length = 2
	while True:
		rows = rows[skip_length:]
		script_name = rows[mid_index].find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
		ActionChains(driver).move_to_element(script_name).click().send_keys(Keys.PAGE_DOWN).perform()
		time.sleep(.01)
		skip_length += 1
		main_table = driver.find_element(By.XPATH, main_table_xpath)
		rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
		if not int(rows[0].get_attribute("data-row-index")) == 0:
			break

	rows = rows[skip_length:]
	row = rows[mid_index]
	script_name = row.find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
	ActionChains(driver).move_to_element(script_name).click().perform()

	for _ in range(int(row.get_attribute("data-row-index"))):
		ActionChains(driver).send_keys(Keys.ARROW_UP).perform()
		time.sleep(.05)

	print("\033[92m" + " Done." + "\033[0m")
	return num_lines-1,mid_index-1,skip_length

def register_fingerprint(fingerprint):
	global processed_entries
	global entries_cache_path
	processed_entries.add(fingerprint)
	with open(entries_cache_path, "wb") as f:
		pk.dump(processed_entries, f)

def register_row_number(row_number):
	global processed_rows
	global rows_cache_path
	processed_rows.add(row_number)
	with open(rows_cache_path, "wb") as f:
		pk.dump(processed_rows, f)

def parse_date(date_str):
	date = None
	for date_format in ['%b %d', '%b %d, %Y', '%d %b', '%d %b, %Y', '%m %d', '%m %d, %Y', '%d %m', '%d %m, %Y']:
		try:
			date = datetime.strptime(date_str, date_format)
			if date.year == 1900:
				current_year = datetime.now().year
				date = date.replace(year=current_year)
			return date
		except ValueError:
			continue
	raise InvalidDateFormatException(f"Invalid date format: {date_str}")

def log_not_found(script_name):
	print("\033[91m" + "Files missing." + "\033[0m")
	with open('missing_logs.txt','a') as f:
		f.write(f"{script_name}\n")

def new_tab_main(script_name,files):
	global driver
	try:
		print("Waiting for page to load")
		while True:
			try:
				attachments_header = WebDriverWait(driver, 60).until(
					EC.presence_of_element_located(
						(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div/h2")
					)
				)
				break
			except TimeoutException:
				driver.refresh()

		attachment_texts = []
		try:
			attachments_list = WebDriverWait(driver, 2).until(
				EC.presence_of_element_located(
					(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div/div/div/div/div[2]/div/div/div/div/div")
				)
			)

			time.sleep(.1)

			attachment_divs = attachments_list.find_elements(
				By.XPATH, ".//div[@role='presentation' and contains(@class, 'ms-List-cell') and @data-list-index]"
			)

			for div in attachment_divs:
				link_element = div.find_element(By.XPATH, ".//a[@role='link' and @tabindex='-1']")
				attachment_texts.append(link_element.text)

		except TimeoutException:
			pass
		except NoSuchElementException:
			pass

		start_text = attachments_header.text

		filtered_files = [f for f in files if not any((text in f or f in text) for text in attachment_texts)]

		if filtered_files:

			add_attachment_button = WebDriverWait(driver, 60).until(
				EC.presence_of_element_located(
					(By.XPATH, "//li[@command='add-attachment']")
				)
			)
			add_attachment_button.click()

			print("Waiting for input box to appear")
			WebDriverWait(driver, 60).until(
				EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div/div/div/div[2]/div/div[2]/div[1]/div/div[1]/div/div[1]/span"))
			)

			file_input = WebDriverWait(driver, 60).until(
				EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
			)

			for file_path in filtered_files:
				file_input.send_keys(file_path)

			ok_button = WebDriverWait(driver, 60).until(
				EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div/div/div/div[2]/div/div[2]/div[2]/div/span[1]/button[@data-is-focusable='true']"))
			)
			time.sleep(.1)
			ActionChains(driver).move_to_element(ok_button).click().perform()

			print("Waiting for attachment count to update")
			while True:
				time.sleep(.5)
				attachments_header = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div/h2")
				if not attachments_header.text == start_text:
					break
			message = "\033[92m" + "Files attached successfully." + "\033[0m"
		else:
			message = "\033[92m" + "Files already attached." + "\033[0m"

		attachment_texts = []
		try:
			attachments_list = WebDriverWait(driver, 2).until(
				EC.presence_of_element_located(
					(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div/div/div/div/div[2]/div/div/div/div/div")
				)
			)

			time.sleep(.1)

			attachment_divs = attachments_list.find_elements(
				By.XPATH, ".//div[@role='presentation' and contains(@class, 'ms-List-cell') and @data-list-index]"
			)

			for div in attachment_divs:
				link_element = div.find_element(By.XPATH, ".//a[@role='link' and @tabindex='-1']")
				attachment_texts.append(link_element.text)

		except TimeoutException:
			pass
		except NoSuchElementException:
			pass

		if (not any(text.endswith("_.lst") for text in attachment_texts)) or (not any(text.endswith(".lst") for text in attachment_texts)):
			log_not_found(script_name)
			return False
		else:
			print(message)
			return True
	except:
		log_not_found(script_name)
		return False

def process_row(row):
	global driver
	global processed_rows
	global processed_entries
	global original_tab
	try:

		data_row_index =int(row.get_attribute("data-row-index"))

		if data_row_index in processed_rows:
			print("\033[33m" + f"\nrow {data_row_index+1:>4} : \033[92mALREADY SEEN" + "\033[0m")
			return

		tds = row.find_elements(By.TAG_NAME, "td")

		script_name = tds[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
		print("\033[33m" + f"\nrow {data_row_index+1:>4} : {script_name.text}" + "\033[0m")

		text = tds[4].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]").text.lower()

		print("Assessing passed status")
		if not text == "passed":
			print("\033[91m" + "Not passed." + "\033[0m")
			register_row_number(data_row_index)
			return

		print("Waiting for entries table to appear")
		ActionChains(driver).move_to_element(script_name).double_click().perform()
		while True:
			try:
				new_table = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, new_table_xpath)))
				break
			except TimeoutException:
				pass
		time.sleep(.1)
		new_rows = new_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and @data-row-index]")[:6]
		new_rows = [
			(row,[span.text for span in row.find_elements(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")]) for row in new_rows
		]
		print("Waiting for entries table to sort")
		new_rows = sorted(new_rows, key=lambda r: parse_date(r[1][1]), reverse=True)

		valid_prossessed = True

		script_name = script_name.text
		for new_row,texts in new_rows:
			fingerprint = hashlib.sha256(str([script_name]+texts).encode('utf-8')).hexdigest()

			if fingerprint in processed_entries:
				print("\033[92m" + "Entry already seen." + "\033[0m")
				register_fingerprint(fingerprint)
				break
			elif texts and texts[0] == "Passed":

				configuration = texts[2].replace('/','')

				files = []
				if (not (configs_to_files.get(configuration,None) == None)) and (not (configs_to_files[configuration].get(script_name,None) == None)):
					files = configs_to_files[configuration][script_name]

				print("Waiting for new tab to open")

				existing_handles = set(driver.window_handles)
				while True:
					driver.switch_to.window(original_tab)

					ActionChains(driver).move_to_element(new_row).double_click().perform()
					time.sleep(3)

					updated_handles = set(driver.window_handles)
					new_handles = updated_handles - existing_handles

					if len(new_handles) == 1:
						break

					while new_handles:
						new_handle = new_handles.pop()
						driver.switch_to.window(new_handle)
						driver.close()

				new_handle = new_handles.pop()

				driver.switch_to.window(new_handle)
				time.sleep(4)

				valid_prossessed = new_tab_main(script_name,files)

				driver.close()
				driver.switch_to.window(driver.window_handles[0])

				register_fingerprint(fingerprint)
				break

			register_fingerprint(fingerprint)

		WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, close_button_xpath)))
		close_button = driver.find_element(By.XPATH, close_button_xpath)

		ActionChains(driver).move_to_element(close_button).click().perform()

		if valid_prossessed:
			register_row_number(data_row_index)

	except StaleElementReferenceException:
		pass

def main():
	global driver
	global original_tab

	populate_configs_to_files()
	fetch_cached_files()
	start_driver()

	driver.get("https://global.tfs.landisgyr.net/tfs/DefaultCollection_Brazil/HQA/_testPlans")

	input("\033[33m" + """
Please follow the steps below; then come back and press ENTER to continue: \033[0m
1. Sign in into ADO with your credentials.
2. Select the desired test suite.
3. Wait for the list of test points to appear. 

\033[33mPress ENTER to start the script: """ + "\033[0m")

	original_tab = driver.current_window_handle

	try:
		WebDriverWait(driver, 1).until(
			EC.presence_of_element_located((By.XPATH, loading_dots_xpath))
		)
		print("\033[33m" + f"\nWaiting for table to completely load." + "\033[0m",end="")
		while True:
			try:
				WebDriverWait(driver, 120).until_not(
					EC.presence_of_element_located((By.XPATH, loading_dots_xpath))
				)
				break
			except TimeoutException:
				pass
		print("\033[92m" + " Done." + "\033[0m")
	except TimeoutException:
		pass

	try:
		WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, main_table_xpath)))
	except:
		print("\033[91m" + "\nCould not find Test Points list in the page. Aborting." + "\033[0m")
		driver.quit()
		sys.exit(1)

	print("""
From this point onward, you won't be allowed to move or resize the window.
Doing so may crash the script. Please make sure the window is where you 
want it to be and then either press ENTER to continue or CTR+C twice to
quit the script.
""")
	input("\033[33m" + f"Continue: " + "\033[0m")

	num_lines,mid_index,skip_length = assess_table_parameters()

	num = min(num_lines,skip_length)-1

	main_table = driver.find_element(By.XPATH, main_table_xpath)
	rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
	rows = rows[:num]

	for row in rows:
		process_row(row)

	print("\033[33m" + f"\nScroling down." + "\033[0m")
	for i in range(num):
		main_table = driver.find_element(By.XPATH, main_table_xpath)
		rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
		if i+mid_index >= len(rows):
			print("\033[33m" + f"\nEnd of list reached. Quitting." + "\033[0m")
			driver.quit()
			return 
		row = rows[i+mid_index]
		script_name = row.find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
		ActionChains(driver).move_to_element(script_name).click().send_keys(Keys.PAGE_DOWN).perform()
		time.sleep(.1)

	main_table = driver.find_element(By.XPATH, main_table_xpath)
	rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
	rows = rows[num:skip_length]

	if rows:
		for row in rows:
			process_row(row)

		print("\033[33m" + f"\nScroling down." + "\033[0m")
		for i in range(len(rows)):
			main_table = driver.find_element(By.XPATH, main_table_xpath)
			rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
			if i+mid_index >= len(rows):
				print("\033[33m" + f"\nEnd of list reached. Quitting." + "\033[0m")
				driver.quit()
				return 
			row = rows[i+mid_index]
			script_name = row.find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
			ActionChains(driver).move_to_element(script_name).click().send_keys(Keys.PAGE_DOWN).perform()
			time.sleep(.1)

	while True:
		main_table = driver.find_element(By.XPATH, main_table_xpath)
		rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")

		rows = rows[skip_length:]
		rows = rows[:num_lines]

		for row in rows:
			process_row(row)

		print("\033[33m" + f"\nScroling down." + "\033[0m")
		for _ in range(num_lines):
			main_table = driver.find_element(By.XPATH, main_table_xpath)
			rows = main_table.find_elements(By.XPATH, ".//tbody//tr[contains(@class, 'bolt-table-row') and contains(@class, 'bolt-list-row') and contains(@role, 'row') and @data-row-index]")
			rows = rows[skip_length:]
			if mid_index >= len(rows):
				print("\033[33m" + f"\nEnd of list reached. Quitting." + "\033[0m")
				driver.quit()
				return 
			script_name = rows[mid_index].find_elements(By.TAG_NAME, "td")[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")
			ActionChains(driver).move_to_element(script_name).click().send_keys(Keys.PAGE_DOWN).perform()
			time.sleep(.1)



if __name__=="__main__":
	main()

# MoveTargetOutOfBoundsException