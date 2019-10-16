# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import smartsheet
import logging
import os.path
import time

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = None

_dir = os.path.dirname(os.path.abspath(__file__))

def update_project_sheet(sheet_id):
    sheet = smart.Sheets.get_sheet(sheet_id)
    print("Found " + sheet.name + " @ " + sheet.permalink + "!")
    print("Doing some work to this sheet...")
    print("")
    time.sleep(3)

    # TODO: Do desired work to each project's sheet here!

print("Starting ...")

# Initialize client
smart = smartsheet.Smartsheet(access_token)

# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Load entire sheet
summary_sheet_id = 1272282845865860
summary_sheet = smart.Sheets.get_sheet(summary_sheet_id)

print("Loaded sheet: " + summary_sheet.name)
print(summary_sheet.permalink)
print("")

# The API identifies columns by Id, but it's more convenient to refer to column names.
# Build column map for later reference - translates column names to column id
summary_column_map = dict((column.title, column.id) for column in summary_sheet.columns)


# Check each row on the summary sheet
for row in summary_sheet.rows:

    # Get the cell in the Project Number column to find the Metrics Sheet for each project.
    # ie: Follow the cell in YTD Actual to find Budget sheet.
    project_number_cell = row.get_column(summary_column_map['Project Number'])


    # Only project rows have cell links - inherently skips the headers since there won't be a cell link
    if project_number_cell.link_in_from_cell:
        project_name_cell = row.get_column(summary_column_map['Project Name'])
        print('Found project "' + project_name_cell.value + '" on summary row #' + str(row.row_number))
        source_sheet_id = project_number_cell.link_in_from_cell.sheet_id
        update_project_sheet(source_sheet_id)
    else:
        continue


print("Done")
