import pandas as pd
import pyodbc

def read_and_clean_excel(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Remove blank columns
    df.dropna(axis=1, how='all', inplace=True)
    
    # Convert the Date column to a proper date format
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    return df

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

    columns_with_types = "[Date] DATE, " + ", ".join([f"[{col}] NVARCHAR(MAX)" for col in columns[1:]])
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

# Example usage
file_path = 'C:\\Users\\nayan.khemka\\Downloads\\scraped_data.xlsx'  # Your provided file path
df = read_and_clean_excel(file_path)
save_table_to_db(df)
