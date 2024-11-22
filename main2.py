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
processed_rows = set()
astray = None
original_tab = None

main_table_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div/div/div/div[2]/table"
loading_dots_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div/div/div/div[1]/div[1]/span[2]"
new_table_xpath = "/html/body/div[2]/div/div/div/div/div/div[3]/div/div/div/table"
close_button_xpath = "/html/body/div[2]/div/div/div/div/div/div[2]/button"

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

def new_tab_main(script_name):
	global driver
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

	if (not script_name+"_.lst" in attachment_texts) or (not script_name+".lst" in attachment_texts):
		input("\033[92m" + f"Press ENTER to continue: " + "\033[0m")
	else:
		print("\033[92m" + "Files already attached." + "\033[0m")


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

def process_row(row):
	global driver
	global processed_rows
	global astray
	global original_tab

	try:

		data_row_index =int(row.get_attribute("data-row-index"))

		if data_row_index < 2176 or data_row_index in processed_rows:
			print("\033[33m" + f"\nrow {data_row_index+1:>4} : \033[92mALREADY SEEN" + "\033[0m")
			return

		processed_rows.add(data_row_index)

		tds = row.find_elements(By.TAG_NAME, "td")

		script_name = tds[2].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]")

		if not script_name.text in astray:
			print("\033[33m" + f"\nrow {data_row_index+1:>4} : \033[92mALREADY SEEN" + "\033[0m")
			return

		print("\033[33m" + f"\nrow {data_row_index+1:>4} : {script_name.text}" + "\033[0m")

		text = tds[4].find_element(By.XPATH, ".//span[contains(@class, 'text-ellipsis') and contains(@class, 'body-m')]").text.lower()

		print("Assessing passed status")
		if not text == "passed":
			print("\033[91m" + "Not passed." + "\033[0m")
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

		script_name=script_name.text
		for new_row,texts in new_rows:
			if texts and texts[0] == "Passed":

				configuration = texts[2].replace('/','')

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

				new_tab_main(script_name)

				# astray.discard(script_name.text)
				# with open("missing_logs.txt", 'w') as f:
				#     f.writelines(f"{item}\n" for item in astray)

				driver.close()
				driver.switch_to.window(driver.window_handles[0])

				break

		WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, close_button_xpath)))
		close_button = driver.find_element(By.XPATH, close_button_xpath)

		ActionChains(driver).move_to_element(close_button).click().perform()

	except StaleElementReferenceException:
		pass

def main():
	global driver
	global astray
	global original_tab

	with open("missing_logs.txt",'r') as f:
		astray = set(line[:-1] for line in f.readlines())
	
	start_driver()

	driver.get("https://global.tfs.landisgyr.net/tfs/DefaultCollection_Brazil/HQA/_testPlans")

	input("\033[33m" + "Press ENTER to start the script: " + "\033[0m")

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