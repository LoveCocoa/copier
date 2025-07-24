import streamlit as st
import pandas as pd
import numpy as np
from datetime import timedelta
import calendar
from io import BytesIO
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
import re

def date_to_week(date):
    dt = pd.to_datetime(date).to_pydatetime()
    weekday = dt.weekday()
    days_since_friday = (weekday - 4) % 7
    friday = dt - timedelta(days=days_since_friday)
    week_start = friday
    week_end = friday + timedelta(days=6)
    week_start_str = f"{week_start.day:02d} {calendar.month_name[week_start.month]}"
    week_end_str = f"{week_end.day:02d} {calendar.month_name[week_end.month]}"
    return f"{week_start_str} - {week_end_str}"
#insert data dictionary for functional location code
code_to_text = {
    "901":"Propulsion System",
    "902":"Control Circuit",
    "903":"Auxiliary Power Supply",
    "904":"ATP/AO Signalling",
    "905":"Lighting",
    "906":"Air Conditioning and Ventilation System",
    "907":"Onboard Communication System",
    "908":"Doors",
    "909":"Special Equipment",
    "911":"Carbody Shell",
    "912":"Carbody Interior",
    "913":"Carbody Exterior",
    "915":"Bogies",
    "916":"Brake System",
    "918":"Coupling and Interconnection",
    "922":"Train Management System(TMS)",
    "923":"Liquid Cooling System"}
# Function to get code text based on the code
def get_code_text(code):
    # if the code has less than 17 characters, return "Rolling Stock"
    if len(code) > 17:
    #if the 17th character to 20th character is in the dictionary, return the corresponding text
        code_text = code_to_text.get(code[16:19], "#N/A")
        return code_text
    else:
        return "Rolling Stock"

rules = {
    'Guide tire worn out': ['Guide tire', 'worn'],  # Must contain ALL words
    'Guide tire warning': ['Guide tire', 'warning'],
    'Load tire worn out':['load','worn'],
    'Load tire warning':['load','warning'],
    'CCD worn out': ['CCD', 'Worn out'],
    'Train not responding in TWP': ['TWP'],
    'Smoke alarm': ['Smoke alarm'],
    'APU faulty': ['APU faulty'],
    'Guide tire lost signal': ['Guide tire', 'signal'],
    'CCD replacement': ['Collector'],
    'CCD crack': ['CCD'],
    'Brake pressure low':['TWP','Brake'],
    'Ceiling loud noise':['ceiling','sound'],
    'Door major failure':['door major'],
    'Gangway lound noise':['gangway loud'],
    'Liquid cooling system':['liquid'],
    'MTC Major failure':['MTC','major','failure'],
    'Steering Cylinder leak':['steering','leak'],
    'Steering Cylinder warning':['steering','warning'],
    'Water dripping':['water','drip'],
    'Wheel Arc':['Allcar','arc'],
    'Wheel well door lock braket':['bracket']
    }
# Pre-compile regex patterns for speed
compiled_patterns = {}
for category, keywords in rules.items():
    # Regex: (?=.*word1)(?=.*word2)(?=.*word3)
    pattern = ''.join(f'(?=.*{keyword})' for keyword in keywords)
    compiled_patterns[category] = re.compile(pattern, flags=re.IGNORECASE)

# Apply classification
def classify_text(text):
    if pd.isna(text):
        return None
    text = str(text).lower()
    for category, pattern in compiled_patterns.items():
        if pattern.search(text):
            return category
    return None  # If no match

def process_excel(df):
    # Processing logic (unchanged)
    s = df["Location"].astype(str).str
    last_two = s[-2:]
    is_sw = df["Description"].str.contains('SW', case=False, na=False)
    df["System"] = np.where(is_sw, "Track switch", "YM" + last_two)
    df["Week"] = df["Malfunction Start"].apply(date_to_week)
    df["Problem"] = df["Description"].apply(classify_text)
    df["Malfunction Start"] = pd.to_datetime(df["Malfunction Start"]).dt.strftime("%d/%m/%Y")
    df["Malfunction End"] = pd.to_datetime(df["Malfunction End"]).dt.strftime("%d/%m/%Y")
    df["Sub-system - Functional location"] = df["Functional Location"].apply(get_code_text)
    # Column reordering
    cols = df.columns.tolist()
    cols.insert(5, cols.pop(cols.index("System")))
    cols.insert(11, cols.pop(cols.index("Week")))
    cols.insert(6, cols.pop(cols.index("Sub-system - Functional location")))
    cols.insert(7, cols.pop(cols.index("Problem")))
    df = df[cols]
    return df

def main():
    st.set_page_config(page_title="Excel Processor", layout="wide")
    st.title("📊 Excel File Processor")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Read and process file
            df = pd.read_excel(uploaded_file)
            processed_df = process_excel(df)
            
            # Show preview
            st.success("File processed successfully!")
            st.subheader("Preview of Processed Data")
            st.dataframe(processed_df.head(), use_container_width=True)
            
            # Create formatted Excel output
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_df.to_excel(writer, index=False, sheet_name="Processed Data")
                
                # Access the workbook and worksheet for formatting
                workbook = writer.book
                worksheet = writer.sheets["Processed Data"]
                
                # Define the table range
                max_row, max_col = processed_df.shape
                table_range = f"A1:{chr(64 + max_col)}{max_row + 1}"
                
                # Create table with proper syntax
                tab = openpyxl.worksheet.table.Table(displayName="ProcessedTable", ref=table_range)
                
                # Add style to the table
                style = openpyxl.worksheet.table.TableStyleInfo(
                    name="TableStyleMedium9",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                tab.tableStyleInfo = style
                
                # Add the table to the worksheet
                worksheet.add_table(tab)
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Download button
            st.download_button(
                label="⬇️ Download Processed File",
                data=output.getvalue(),
                file_name=f"processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.error("Please ensure your file contains the required columns")

if __name__ == "__main__":
    main()