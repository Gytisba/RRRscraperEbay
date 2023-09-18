import time
import csv
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# Set up Firefox options
firefox_options = Options()
firefox_options.binary_location = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"

# Specify path to the geckodriver
gecko_driver_path = "./geckodriver.exe"
s = Service(gecko_driver_path, log_output="geckodriver.log")

# Create a new instance of the Firefox browser
driver = webdriver.Firefox(service=s, options=firefox_options)
policy_accepted = False

def accept_rrr_privacy_policy():
    global policy_accepted

    # If the policy has already been accepted, just return
    if policy_accepted:
        return

    try:
        wait = WebDriverWait(driver, 10)
        # Wait for the privacy policy button to appear and then click it
        wait.until(EC.element_to_be_clickable((By.ID, 'CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll'))).click()
        policy_accepted = True
    except TimeoutException:
        print("Privacy policy button did not appear or was not found on rrr.lt.")



# ... [your imports and browser setup here]

def scrape_rrr(page_number):
    base_url = f'https://rrr.lt/paieska?man_id=1&cmc=6&cm=1391&mfi=1,6,1391;&q=&page={page_number}'
    driver.get(base_url)
    accept_rrr_privacy_policy()
    time.sleep(2)
    items = driver.find_elements(By.CLASS_NAME, "products__box")
    data = []

    for index, item in enumerate(items, start=1):  # Start indexing from 1
        code, product_price = None, None

        try:
            # Dynamically build the XPath for product code using the index
            code_xpath = f"/html/body/div[1]/main/div[2]/section[1]/section[2]/div[{index}]/div[1]/p[1]/a[1]"
            code_element = driver.find_element(By.XPATH, code_xpath)
            code = code_element.text.strip()
        except:
            print(f"Error processing item number {index}: Code not found on page {page_number}")
            print(code_xpath)
        try:
            # Dynamically build the XPath for product price using the index
            price_xpath = f"/html/body/div[1]/main/div[2]/section[1]/section[2]/div[{index}]/div[2]/div/strong"
            price_element = driver.find_element(By.XPATH, price_xpath)
            product_price_text = price_element.text.replace("€", "").replace(",", ".").strip()
            product_price = float(product_price_text)
        except:
            print(f"Error processing item number {index}: Product price not found on page {page_number}")

        # Only add data if we have code and product price
        if code and product_price is not None:
            data.append((code, product_price))

    return data

def scrape_ebay(item_code):
    driver.get(f'https://www.ebay.com/sch/i.html?_from=R40&_trksid=p2334524.m570.l1313&_nkw={item_code}&_sacat=0&LH_TitleDesc=0&_odkw=7065702&_osacat=0&LH_Complete=1&LH_Sold=1')
    time.sleep(3)  # give the page some time to load
    
    try:
        element_present = EC.presence_of_element_located((By.CLASS_NAME, 's-item__price'))
        WebDriverWait(driver, 10).until(element_present)
    except:
        print("Timed out waiting for eBay page to load")
        return []
    
    prices = []
    
    try:
        # Extracting all price elements using class name
        price_elements = driver.find_elements(By.CLASS_NAME, 's-item__price')
        
        # Extracting text from each price element, stripping extra characters and converting to float
        for price_elem in price_elements:
            price_text = price_elem.text.replace("$", "").replace(",", "")
            if price_text:  # Check if price text is not empty
                prices.append(float(price_text))
        
    except Exception as e:
        print(f"Error while extracting prices: {e}")
    
    return prices



def main():
    all_codes_prices = []

    # Loop through 5 pages to collect all the codes and prices
    for page in range(1, 3):
        all_codes_prices.extend(scrape_rrr(page))

    # Create a list to store the rows of data for the DataFrame
    data_rows = []

    for code, price in all_codes_prices:
        ebay_prices = scrape_ebay(code)
        if ebay_prices:
            lowest_ebay_price = min(ebay_prices)
            diff = price - lowest_ebay_price
            data_rows.append([code, price, lowest_ebay_price, diff])
        else:
            data_rows.append([code, price, "Not Found on eBay", "N/A"])

    # Create a pandas DataFrame
    df = pd.DataFrame(data_rows, columns=["Code", "RRR Price", "Lowest eBay Price", "Difference"])

    # Save DataFrame to Excel
    with pd.ExcelWriter('results.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        # Get the xlsxwriter objects
        worksheet = writer.sheets['Sheet1']

        # Add some basic formatting
        header_format = writer.book.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # Write the column headers with the defined format.
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Adjust the column width
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max(series.astype(str).apply(len).max(),  # max length in column
                          len(str(series.name)))  # length of column name/header
            worksheet.set_column(idx, idx, max_len)  # set column width

    driver.close()

if __name__ == "__main__":
    main()
