import sqlite3
import pandas as pd

# Path to the database
db_path = "/Users/johnathenwigfall/Desktop/Apple customer Feedback/feedback.db"
conn = sqlite3.connect(db_path)

# Create a cursor object to interact with the database
cursor = conn.cursor()

# List all tables
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
print("Tables:", tables)

# Dictionary to store customer feedback by table
customer_feedback_dict = {}

# Fetch rows and columns from each table and store them in a dictionary
for table in tables:
    table_name = table[0]  # Extract the table name
    print(f"\nFetching rows from {table_name}:")
    
    # Fetch all rows and column names from the table
    cursor.execute(f"SELECT * FROM {table_name};")
    rows = cursor.fetchall()
    columns = [description[0] for description in cursor.description]  # Get column names
    
    # Store the data as a DataFrame
    customer_feedback_dict[table_name] = pd.DataFrame(rows, columns=columns)

# Close the connection
conn.close()

# Save each table's DataFrame to an Excel file
excel_path = "/Users/johnathenwigfall/Desktop/Apple customer Feedback/Apple Users Feedback.xlsx"
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for table_name, df in customer_feedback_dict.items():
        df.to_excel(writer, sheet_name=table_name, index=False)

print(f"Customer feedback has been saved to {excel_path}")
