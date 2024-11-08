import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook

# Load the Excel file
file_path = "./law_sheet.xlsx"
df = pd.read_excel(file_path)

# Setup Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver_service = Service('./chromedriver.exe')  

driver = webdriver.Chrome(service=driver_service, options=chrome_options)

# for index, row in df.iterrows():
for index, row in df.iloc[281:288].iterrows():
    
    # law_link = row['G']
    # law_name = row['B'].replace('.pdf', '')
    law_link = row['링크']
    law_name = row['파일명'].replace('.pdf', '')

    # Connect to the link using Selenium
    driver.get(law_link)
    time.sleep(1)  # Wait for the site to load

    try:
        driver.switch_to.frame("lawService")  # Replace with the iframe ID or use another method to locate it

        # Check if the iframe contains nested iframes and switch accordingly
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if len(iframes) > 0:
            # Switch to the first nested iframe, if applicable
            driver.switch_to.frame(iframes[0])

    except:
        print("Failed to switch to iframe. Check if the iframe ID or locator is correct.")

    try:
        # Get text information from #conTop > h2 > span
        span_text = driver.find_element(By.CSS_SELECTOR, "#conTop > h2 > span").text
        span_text = span_text[span_text.find(":")+2:span_text.find(")")-1]

        # df.at[index, 'C'] = span_text
        df.at[index, '소방법령(약칭)'] = span_text
    except:
        # If no span tag is found, copy the value from column D
        # df.at[index, 'C'] = row['D']
        df.at[index, '소방법령(약칭)'] = row['소방법령']

    # Create a directory with the name from column C if it doesn't exist
    directory_name = df.at[index, '소방법령(약칭)']
    if not os.path.exists(directory_name):
        os.makedirs(directory_name)

    # Create an Excel file in the folder with the name of column B (without ".pdf")
    excel_file_path = os.path.join(directory_name, f"{law_name}.xlsx")
    workbook = Workbook()
    main_sheet = workbook.active
    main_sheet.title = "Main"
    main_sheet['A1'] = law_link

    has_span_a_element = False

    try:
        # Connect to the link using Selenium
        driver.get(law_link)
        time.sleep(1)  # Wait for the site to load

        try:
            driver.switch_to.frame("lawService")  # Replace with the iframe ID or use another method to locate it

            # Check if the iframe contains nested iframes and switch accordingly
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if len(iframes) > 0:
                # Switch to the first nested iframe, if applicable
                driver.switch_to.frame(iframes[0])

        except:
            print("Failed to switch to iframe. Check if the iframe ID or locator is correct.")

        # Retrieve information from #conScroll > ul
        ul_element = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#conScroll > ul"))
        )
        li_elements = ul_element.find_elements(By.TAG_NAME, "li")
        
        for li in li_elements:
            try:
                sheet_name = li.find_element(By.CSS_SELECTOR, "span > a:nth-child(3)").text
                if not '별표' in sheet_name:
                    continue
                sheet_name = sheet_name[sheet_name.find("[")+1:sheet_name.find("]")].replace(" ", "")
                # Create a new sheet in the workbook with the cleaned-up sheet name
                workbook.create_sheet(title=sheet_name)
                has_span_a_element = True  # Mark that a span > a:nth-child(3) was found
            except:
                continue

    except TimeoutException:
        print("Timed out waiting for the ul element to load.")
        # pass

    # Save the workbook
    workbook.save(excel_file_path)

    # If no span > a:nth-child(3) element was found, rename the file by adding "[별표없음]"
    if not has_span_a_element:
        if os.path.exists(excel_file_path):
            new_excel_file_path = os.path.join(directory_name, f"[별표없음]{law_name}.xlsx")
            os.rename(excel_file_path, new_excel_file_path)
            excel_file_path = new_excel_file_path  # Update the path for consistency
        else:
            print(f"File not found: {excel_file_path}")


# Save the updated dataframe back to the Excel file
df.to_excel("updated_law_sheet.xlsx", index=False)

# Close the Selenium WebDriver
driver.quit()
