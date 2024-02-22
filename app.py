import os
import openai
import gspread
from google.oauth2 import service_account

# Setup OpenAI
openai.api_key ="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"

# Setup Google Sheets
spread_sheets_id = "1_riJ6THApv06jc86OFauxcbMi0yenf66jQTJlhV6qMY" 
scope = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = service_account.Credentials.from_service_account_file(
    "C:/Users/abdes/Desktop/key.json.json", scopes=scope
)
gc = gspread.authorize(credentials)
sheet = gc.open_by_key(spread_sheets_id).sheet1

def classify_comment(comment):
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": """
                    You are a helpful assistant designed to classify comments into:
                    
                    Neutral, Request, Question, Appreciation, Error or Other.
                    
                    Must respond with one word.
                    
                    Comment: "{comment}"
                    
                    Category:
                """
                },
                {"role": "user", "content": comment} 
            ],
        )
    

        category = response.choices[0].message.content
        print(f"Comment: '{comment}' -> Category: '{category}'")
        return category.title()
    except Exception as e:
        print(f"Error in classifying comment: {e}")
        return "Process Failed"

# Read data from the sheet
data = sheet.get_all_values()
header = data[0]  # Assuming the first row is the header
comments_rows = data[1:]  # Exclude header

# Iterate over each row that contains comments
for i, row in enumerate(comments_rows, start=2):  # Google Sheets is 1-indexed; header is row 1, so data starts at row 2
    comment = row[0]  # Assuming the comment is in the second column
    category = classify_comment(comment)
    # Update the 3rd column (index 2) with the classification
    sheet.update_cell(i, 3, category)

print("The Google Sheet has been updated with categories for each comment.")
