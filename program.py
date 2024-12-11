from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

def scrape_table():
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--disable-dev-shm-usage') 
    options.add_argument('--no-sandbox')

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    url = "https://www.gasnetworks.ie/corporate/gas-regulation/transparency-and-publicat/dashboard-reporting/balancing-actions-and-prices/imbalance-prices/"
    driver.get(url)

    # Wait for the table to be present
    table = driver.find_element(By.CSS_SELECTOR, 'table')  # Adjust the selector as needed
    rows = table.find_elements(By.TAG_NAME, 'tr')

    data = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        if cols == []:
            cols = row.find_elements(By.TAG_NAME, 'th')
        data.append([col.text for col in cols])
    driver.quit()
    
    # Convert the data to a DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])  # Assuming first row contains column names
    # df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    print(df)
    return df

def save_table_to_excel(data, file_name='scraped_data.xlsx'):
    data.to_excel(file_name, index=False)
    print(f'Data saved to {file_name}')

# Example usage
table_data = scrape_table()
save_table_to_excel(table_data)
