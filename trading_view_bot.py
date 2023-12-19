import os
import time
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import pandas as pd
from io import StringIO


urls = [
    "https://www.tradingview.com/markets/stocks-usa/market-movers-large-cap/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-active/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-gainers/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-losers/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-most-volatile/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-overbought/",
    "https://www.tradingview.com/markets/stocks-usa/market-movers-oversold/",   
]


driver_path = os.path.join(os.getcwd(), "chromedriver-win64", "chromedriver.exe")
chrome_service = Service(driver_path)
chrome_options = Options()

browser = Chrome(service=chrome_service, options=chrome_options)
browser.implicitly_wait(7)
browser.maximize_window()

for url in urls:
    browser.get(url)

    # Step 1.1 Getting the file base name
    file_base_name = url.split("/")[-2]

    print(f"Scrapping {file_base_name}...")

    # Step 1.2 Create an Excel writer's object 
    xlwritter = pd.ExcelWriter(f"{file_base_name}.xlsx", engine="xlsxwriter")


    # Step 2.1 iterate each report/category
    categories = [
        "Overview", "Performance", "Valuation", "Dividends", "Profitability", "Income Statement", 
        "Balance Sheet", "Cash Flow", "Technicals", 
    ]
    i = 1
    for category in categories:
        print(f"Processing report {category}...")
        
        try:
            
            element = browser.find_element(By.XPATH, f"/html/body/div[3]/div[4]/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/button[{i}]/span[1]/span")
            try:
                element.click()
            except ElementNotInteractableException as e:
                print(f"Report {category} is not clickable.")
                continue
                
            # delay execution for the table to load
            time.sleep(3)
            tables = pd.read_html(StringIO(browser.page_source))
            if len(tables) > 1:
                df = tables[1]
                df.replace("-", "", inplace=True)
                df.to_excel(xlwritter, sheet_name=category, index=False)
            else:
                print(f"No table found for report {category}.")
            
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Report {category} is not found.")
            continue
        i += 1
    
    xlwritter._save()
    
browser.quit()







