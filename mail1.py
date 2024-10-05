import pandas as pd

# Load the email formats and disqualified emails
email_df = pd.read_excel('email_formats.xlsx')  # Adjust sheet name if needed
disqualified_df = pd.read_excel('disqualified mail.xlsx')  # Adjust sheet name

# Create a set of disqualified emails for fast lookup
disqualified_emails = set(disqualified_df['disqualified mails'].tolist())

# Function to highlight disqualified emails by changing text color
def highlight_disqualified(email):
    return 'color: red' if email in disqualified_emails else ''

# Apply highlighting to each email format column using map
styled_df = email_df.style.map(highlight_disqualified, 
                                subset=['Format 1', 'Format 2', 'Format 3'])

# Save the styled DataFrame to an Excel file
file_path = 'highlighted_emails.xlsx'
styled_df.to_excel(file_path, engine='openpyxl', index=False)

# Adjust the width of the cells
from openpyxl import load_workbook

# Load the workbook and select the active sheet
wb = load_workbook(file_path)
ws = wb.active

# Set the width of the columns
column_widths = {
    'A': 30,  # Adjust width for Full Name
    'B': 30,  # Adjust width for Format 1
    'C': 30,  # Adjust width for Format 2
    'D': 30   # Adjust width for Format 3
}

for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

# Save the workbook with the adjusted column widths
wb.save(file_path)

print("\nExecution finished. The highlighted emails have been saved to 'highlighted_emails.xlsx'.")

