import pandas as pd

# Define the file path
file_path = '/Users/benprisby/Documents/Sunwhereused-templates3.xlsx'

# Define column names according to your description
column_names = [
    'Template', 'Exp_Mat_Group', 'Make_Item', 'Component', 'Plant',
    'Jan_2024', 'Feb_2024', 'Mar_2024', 'Apr_2024', 'May_2024',
    'Jun_2024', 'Jul_2024', 'Aug_2024', 'Sep_2024', 'Oct_2024',
    'Nov_2024', 'Dec_2024', 'Jan_2025', 'Feb_2025', 'Mar_2025', 'Total'
]

# Read the Excel file
try:
    df = pd.read_excel(file_path, header=0, names=column_names, dtype={
        'Template': str,
        'Exp_Mat_Group': str,
        'Make_Item': str,
        'Component': str,
        'Plant': str,
        'Jan_2024': float,
        'Feb_2024': float,
        'Mar_2024': float,
        'Apr_2024': float,
        'May_2024': float,
        'Jun_2024': float,
        'Jul_2024': float,
        'Aug_2024': float,
        'Sep_2024': float,
        'Oct_2024': float,
        'Nov_2024': float,
        'Dec_2024': float,
        'Jan_2025': float,
        'Feb_2025': float,
        'Mar_2025': float,
        'Total': float
    })

    print("File successfully loaded.")
    print(df.head())

except FileNotFoundError:
    print(f"Error: The file {file_path} was not found.")
except Exception as e:
    print(f"An error occurred: {e}")