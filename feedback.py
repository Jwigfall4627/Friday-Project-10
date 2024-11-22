import sqlite3
import pandas as pd
from textblob import TextBlob

# Path to the database
db_path = "/Users/johnathenwigfall/Desktop/Friday-Project-10/feedback.db"
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
excel_path = "/Users/johnathenwigfall/Desktop/Friday-Project-10/Apple Users Feedback.xlsx"
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for table_name, df in customer_feedback_dict.items():
        df.to_excel(writer, sheet_name=table_name, index=False)

print(f"Customer feedback has been saved to {excel_path}")

import sqlite3
import pandas as pd
from textblob import TextBlob
import matplotlib.pyplot as plt
import seaborn as sns

# Path to the database
db_path = "/Users/johnathenwigfall/Desktop/Friday-Project-10/feedback.db"
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
excel_path = "/Users/johnathenwigfall/Desktop/Friday-Project-10/Apple Users Feedback.xlsx"
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for table_name, df in customer_feedback_dict.items():
        df.to_excel(writer, sheet_name=table_name, index=False)

print(f"Customer feedback has been saved to {excel_path}")

# Load the Excel file
df = pd.read_excel(excel_path, sheet_name=0)  # Load the first sheet for analysis

# Ensure the 'comment' column exists
if 'comment' not in df.columns:
    raise ValueError("The 'comment' column is missing in the Excel file.")

# Define a function for sentiment analysis
def analyze_sentiment(text):
    analysis = TextBlob(text)
    return analysis.sentiment.polarity  # Returns polarity score (-1 to 1)

# Apply the function to the 'comment' column
df['Sentiment Score'] = df['comment'].apply(analyze_sentiment)

# Categorize sentiment based on the score
def categorize_sentiment(score):
    if score > 0:
        return 'Positive'
    elif score < 0:
        return 'Negative'
    else:
        return 'Neutral'

df['Sentiment Category'] = df['Sentiment Score'].apply(categorize_sentiment)

# Print all reviews with their sentiment scores and categories
for index, row in df.iterrows():
    print(f"Review {index + 1}:")
    print(f"Text: {row['comment']}")
    print(f"Sentiment Score: {row['Sentiment Score']}")
    print(f"Category: {row['Sentiment Category']}")
    print("-" * 50)

# Save the results to a new Excel file
output_file = "/Users/johnathenwigfall/Desktop/Friday-Project-10/reviews_with_sentiment.xlsx"
df.to_excel(output_file, index=False)
print(f"Sentiment analysis results saved to {output_file}")

# Data Visualization

# 1. Bar Chart for Sentiment Category Distribution
plt.figure(figsize=(8, 5))
sns.countplot(data=df, x='Sentiment Category', palette='viridis')
plt.title('Sentiment Category Distribution', fontsize=14)
plt.xlabel('Sentiment Category', fontsize=12)
plt.ylabel('Count', fontsize=12)
plt.show()

# 2. Table Summary
summary = df['Sentiment Category'].value_counts().reset_index()
summary.columns = ['Sentiment Category', 'Count']

# Display the table
print("\nSummary Table:")
print(summary)

# Plot the table
plt.figure(figsize=(6, 2))
plt.axis('off')
plt.table(cellText=summary.values, colLabels=summary.columns, cellLoc='center', loc='center')
plt.title('Summary Table of Sentiments', fontsize=14)
plt.show()

# 3. Boxplot for Sentiment Scores
plt.figure(figsize=(8, 5))
sns.boxplot(data=df, x='Sentiment Category', y='Sentiment Score', palette='coolwarm')
plt.title('Sentiment Scores by Category', fontsize=14)
plt.xlabel('Sentiment Category', fontsize=12)
plt.ylabel('Sentiment Score', fontsize=12)
plt.show()

# 4. Pie Chart for Sentiment Distribution
plt.figure(figsize=(6, 6))
df['Sentiment Category'].value_counts().plot.pie(
    autopct='%1.1f%%', startangle=90, colors=sns.color_palette('pastel'))
plt.title('Sentiment Category Proportion', fontsize=14)
plt.ylabel('')
plt.show()
