import os
import pandas as pd
import pdfplumber
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# All files manually downloaded from https://corporate.bwfbadminton.com/players/historical-rankings/?ryear=2015
# A webscraper was attempted but blocked by the website
# Each year is downloaded quarterly (4 files per year, January, April, July, October)
# The week is Week 1 for each month except some years due to limited data, where Week 2 is used as a substitute
#   This is unlikely to cause any major differences in calculations since the rankings are unlikely to vary significantly per week
# Year 2020 happened during COVID, and only January Week 1 (for 2020 Q1) and March Week 3 (for 2020 Q2) were selected due to rankings being frozen (no new data until 2021)


# FILE PATHS
input_folder = r"C:\Users\laivi\OneDrive\Documents\bwf ranking project\World Ranking Data"
output_file = r"C:\Users\laivi\OneDrive\Documents\bwf ranking project\MS_Rankings_Top_200.xlsx"
print("start")

# FUNCTION parse_row
    # Parses each row of rankings
def parse_row(row):
    parts = row.split()  # Split by " " the row into list of components
                         # Some names are 2 words, others 3 or even 4
                         # This is accounted for by working backwards from the list
    try:
        if file_name == "WR 2019-10-01 (Week 40).xlsx":     # for some reason the output from file "WR 2019-10-01 (Week 40).xlsx" is ['200', '86104', 'LOW', 'Pit', 'Seng', 'M', 'AUS', '8460', '17', 'nan', 'nan', 'nan', 'nan']
            parts = parts[0:-4]                             # instead of the standard ['200', '86104', 'LOW', 'Pit', 'Seng', 'M', 'AUS', '8460', '17']                                     
                                                            # if this file is loaded, we will simply do parts = parts[0:-4] to remove the last 4 nan values
        ranking = int(parts[0])
        bwf_id = parts[1]
        gender = parts[-4]
        country = parts[-3]
        points = parts[-2]
        tournaments_played = parts[-1]
        player_name = " ".join(parts[2:-4]) # Because some player names are 3 words instead of 2, splitting will lead to either 2 or 3 pieces
                                            # Therefore for gender, country, points, and tournaments_played, we have to work backwards

        return [ranking, bwf_id, player_name, gender, country, points, tournaments_played]
    except (IndexError, ValueError) as e:
        print(f"Error parsing row: {row} -> {e}")
        return None

# FUNCTION extract_date_from_filename
    # Extracts date from file name because dates are included in file name, so no need to dig through the file content
def extract_date_from_filename(file_path):
    file_name = os.path.basename(file_path)
    try:
        # inputformat "WR 2023-01-03 (Week-01).pdf"
        date_str = file_name.split(" ")[1]  # extract "2023-01-03"
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except Exception as e:
        print(f"Error extracting date from file name: {file_name}, {e}")
        return None

# FUNCTION excel_to_pdf
    # Converts excel files to pdfs
    # Input files are a mix of excel sheets and excel sheets in pdf format
        # However, most files are excel so parsing data through excel sheets rather than pdfs would have been quicker and less computationally intense
        # PDF parser was made first due to the most recent files (also the initial test files) were pdfs instead of excel sheets
        # I am using the pdf parser instead of writing an excel parser since the pdf parser works as intended and the overall computation is still relatively cheap (10-25 min)
            # This analysis is designed as an ad hoc assessment and at maximum re-run on a quarterly basis
def excel_to_pdf(excel_file, pdf_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)
    
    # Create a PDF
    with PdfPages(pdf_file) as pdf:
        fig, ax = plt.subplots(figsize=(12, 8)) # create a figure for the table
        ax.axis('tight')
        ax.axis('off')
        
        # dataframe into table
        table = ax.table(
            cellText=df.values,
            colLabels=df.columns,
            cellLoc='center',
            loc='center'
        )
        
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.auto_set_column_width(col=list(range(len(df.columns))))
        
        pdf.savefig(fig) # save figure to pdf
        plt.close()

# FUNCTION process_pdf_file
    # processes pdf files 
def process_pdf_file(file_path, file_name, data, files_skipped):
    """Process a PDF file and extract relevant rows."""
    ranking_date = extract_date_from_filename(file_path)
    if ranking_date is None:
        print(f"Date not found in file: {file_name}")
        files_skipped.append(file_name)  # skipped file is put into files_skipped list
        return

    with pdfplumber.open(file_path) as pdf:
        row_count = 0  # count number of rows processed
        for page in pdf.pages:
            if row_count >= 200:  # stop processing after 200 rows
                break
            text = page.extract_text() # extract text from each page
            if text:
                lines = text.split("\n")
                for line in lines:
                    if row_count >= 200:  # stop processing after 200 rows
                        break
                    if line.strip() and line[0].isdigit():  # only process rows starting with a digit (Ranking) to ensure we are not processing a header or footer or something
                        parsed_row = parse_row(line)
                        if parsed_row:
                            parsed_row.insert(0, ranking_date)  # add the date to each row
                            data.append(parsed_row)
                            row_count += 1  # increment the count of rows processed by 1


##### EXECUTION #####

# Data Processing
data = []
files_skipped = []
for file_name in os.listdir(input_folder):
    file_path = os.path.join(input_folder, file_name)

    if file_name.endswith(".xlsx"):
        print(f"Converting Excel to PDF: {file_path}")
        # Convert Excel to PDF in-memory
        pdf_file = f"{os.path.splitext(file_name)[0]}.pdf"
        excel_to_pdf(file_path, pdf_file)

        # Process the converted PDF directly
        print(f"Processing converted PDF: {pdf_file}")
        process_pdf_file(pdf_file, file_name, data, files_skipped)

    elif file_name.endswith(".pdf"):
        print(f"Processing PDF: {file_path}")
        process_pdf_file(file_path, file_name, data, files_skipped)


# Process data into dataframe
columns = ["Date", "Ranking", "BWF ID", "Player Name", "Gender", "Country", "Points", "Tournaments Played"]
df = pd.DataFrame(data, columns=columns)

# adjust rows
df["Points"] = pd.to_numeric(df["Points"], errors="coerce") # Convert Points and Tournaments Played to numeric
df["Tournaments Played"] = pd.to_numeric(df["Tournaments Played"], errors="coerce")
df["Points Per Tournament"] = df["Points"] / df["Tournaments Played"] # Calculate Points Per Tournament

# save to Excel
df.to_excel(output_file, index=False)
print(f"Top 200 Men's Singles rankings saved to {output_file}")

# Check if any files were skipped
if files_skipped:
    print("\nFiles Skipped:")
    for skipped_file in files_skipped:
        print(skipped_file)
else:
    print("\nNo files were skipped.")
