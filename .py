from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import pyodbc

def scrape_table():
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--ignore-certificate-errors') 
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
        data.append([col.text for col in cols])
    
    driver.quit()
    
    # Convert the data to a DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])  # Assuming first row contains column names
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# Example usage
table_data = scrape_table()
print(table_data)

def create_connection():
    try:
        connection = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=BNGECO-L-55624;'
            'DATABASE=test;'
            'UID=apple;'
            'PWD=apple'
        )
        print('Connected to SQL Server database')
        return connection
    except pyodbc.Error as e:
        print(f"Error: {e}")
        return None

def create_table_if_not_exists(connection, columns):
    cursor = connection.cursor()

    columns_with_types = "[Date] DATE, " + ", ".join([f"[{col}] DECIMAL(10, 4)" for col in columns[1:]])
    create_table_query = f'''
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='imbalance_prices' AND xtype='U')
    CREATE TABLE imbalance_prices (
        {columns_with_types}
    )
    '''
    cursor.execute(create_table_query)
    connection.commit()
    cursor.close()
    print('Ensured table exists with columns:', columns)

def save_table_to_db(data):
    connection = create_connection()
    if connection:
        create_table_if_not_exists(connection, data.columns)
        
        cursor = connection.cursor()
        
        # Dynamically create the insert query based on the DataFrame columns
        columns = ", ".join([f"[{col}]" for col in data.columns])
        placeholders = ", ".join(["?" for _ in data.columns])
        insert_query = f"INSERT INTO imbalance_prices ({columns}) VALUES ({placeholders})"

        for index, row in data.iterrows():
            cursor.execute(insert_query, tuple(row))
        
        connection.commit()
        cursor.close()
        connection.close()
        print('Data saved to SQL Server database')

table_data = scrape_table()
save_table_to_db(table_data)


