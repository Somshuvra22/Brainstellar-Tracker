import json
import pandas as pd

# Define input files and sheet names
files = {
    "easy.json": "Easy",
    "medium.json": "Medium",
    "hard.json": "Hard",
    "deadly.json": "Deadly"
}

# Status options
status_options = ['To Solve', 'Solved', 'Revisit']

# Create Excel writer
with pd.ExcelWriter("brainstellar_tracker.xlsx", engine="xlsxwriter") as writer:
    for file, sheet in files.items():
        with open(file, "r") as f:
            data = json.load(f)

        for i, p in enumerate(data):
            p["Q.No"] = i + 1
            p["Difficulty"] = sheet
            p["Status"] = "To Solve"
            p["Date Solved"] = ""
            p["Notes"] = ""

        # Create DataFrame and set column order
        df = pd.DataFrame(data)
        df = df[["Q.No", "title", "Difficulty", "url", "Status", "Date Solved", "Notes"]]
        df.columns = ["Q.No", "Title", "Difficulty", "URL", "Status", "Date Solved", "Notes"]

        # Write to Excel
        df.to_excel(writer, sheet_name=sheet, index=False)

        # Get workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets[sheet]
        status_col = df.columns.get_loc("Status")  # zero-indexed

        # Define formats for different statuses
        solved_format = workbook.add_format({
            'bg_color': '#00B050',      # Green background
            'font_color': 'white',
            'bold': True,
            'align': 'center'
        })
        
        to_solve_format = workbook.add_format({
            'bg_color': '#FFC000',      # Amber/Yellow background
            'font_color': 'black',
            'bold': True,
            'align': 'center'
        })
        
        revisit_format = workbook.add_format({
            'bg_color': '#FF0000',      # Red background
            'font_color': 'white',
            'bold': True,
            'align': 'center'
        })

        # Add conditional formatting for Status column
        status_range = f'{chr(65 + status_col)}2:{chr(65 + status_col)}{len(df) + 1}'
        
        # Green for "Solved"
        worksheet.conditional_format(status_range, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': '"Solved"',
            'format': solved_format
        })
        
        # Amber/Yellow for "To Solve"
        worksheet.conditional_format(status_range, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': '"To Solve"',
            'format': to_solve_format
        })
        
        # Red for "Revisit"
        worksheet.conditional_format(status_range, {
            'type': 'cell',
            'criteria': 'equal to',
            'value': '"Revisit"',
            'format': revisit_format
        })

        # Add dropdown for each row in Status column
        for row in range(1, len(df)+1):
            worksheet.data_validation(row, status_col, row, status_col, {
                'validate': 'list',
                'source': status_options,
                'input_message': 'Select a status',
                'error_message': 'Choose from the list only'
            })

        # Optional: Freeze top row and set column widths
        worksheet.freeze_panes(1, 0)
        worksheet.set_column('A:A', 6)   # Q.No
        worksheet.set_column('B:B', 30)  # Title
        worksheet.set_column('C:C', 10)  # Difficulty
        worksheet.set_column('D:D', 50)  # URL
        worksheet.set_column('E:E', 12)  # Status
        worksheet.set_column('F:F', 15)  # Date Solved
        worksheet.set_column('G:G', 30)  # Notes

print("✅ Created brainstellar_tracker.xlsx with color-coded Status column.")
