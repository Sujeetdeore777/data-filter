import pandas as pd
from openpyxl import load_workbook

# Load the email formats and disqualified emails
email_df = pd.read_excel('email_formats.xlsx')  # Adjust sheet name if needed
disqualified_df = pd.read_excel('disqualified mail.xlsx')  # Adjust sheet name

# Create sets of disqualified emails for fast lookup
disqualified_emails = set(disqualified_df['disqualified emails'].dropna().tolist())
disqualified_mails1 = set(disqualified_df['disqualified mails1'].dropna().tolist())

# Combine both sets for complete checks
combined_disqualified_emails = disqualified_emails.union(disqualified_mails1)

# Function to highlight disqualified emails by changing text color
def highlight_disqualified(email):
    return 'color: red' if email in combined_disqualified_emails else ''

# Apply highlighting to each email format column using map
styled_df = email_df.style.map(highlight_disqualified, 
                                subset=['Format 1', 'Format 2', 'Format 3', 'Format 4'])

# Save the styled DataFrame to an Excel file
file_path = 'highlighted_emails.xlsx'
styled_df.to_excel(file_path, engine='openpyxl', index=False)

# Adjust the width of the cells
wb = load_workbook(file_path)
ws = wb.active

# Set the width of the columns
column_widths = {
    'A': 30,  # Adjust width for Full Name
    'B': 30,  # Adjust width for Format 1
    'C': 30,  # Adjust width for Format 2
    'D': 30,  # Adjust width for Format 3
    'E': 30   # Adjust width for Format 4
}

for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

# Save the workbook with the adjusted column widths
wb.save(file_path)

# Print highlighted emails
highlighted_emails = []

# Check each format column for disqualified emails and store them
for column in ['Format 1', 'Format 2', 'Format 3', 'Format 4']:
    for email in email_df[column]:
        if email in combined_disqualified_emails:
            highlighted_emails.append(email)

# Print highlighted emails
print("Highlighted Emails:")
for email in highlighted_emails:
    print(email)

# Print a completion message
print("\nExecution finished. The highlighted emails have been saved to 'highlighted_emails.xlsx'.")
