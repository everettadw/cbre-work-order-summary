import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from datetime import datetime


def generate_max_compliance(source_path, target_path, date=datetime.today()):
    # work order 'Type' filter allowed column values
    work_order_type_filter = [
        "PdM - Work using a Predictive Tool",
        "PM - Preventative Maintenance"
    ]

    # get today's date and format it to use as a filter
    date_filter = date.strftime("%#m/%#d/%Y")

    # put the work order excel file into a dataframe from pandas
    #     and then filter out all work orders except PMs and PdMs
    #     with a max compliance date of today
    df = pd.read_excel(source_path)
    df = df.loc[(df["Type"].isin(work_order_type_filter))
                & (df["PM Compliance Max"] == date_filter)]

    # get rid of useless columns and sort the data in ascending order
    #     (1 - WO Owner) (2 - Equipment) (3 - Description)
    df = df.drop(df.columns[[0, 1, 4, 5, 9, 13, 14, 15, 16]], axis=1)
    df = df.sort_values(by=['1 - WO Owner', 'Equipment', 'Description'])

    # create a new excel file that shows work orders that are going
    #     to be out of compliance by midnight tonight
    with pd.ExcelWriter(
        target_path,
        engine="xlsxwriter",
        engine_kwargs={'options': {'strings_to_numbers': True}}
    ) as writer:
        df.to_excel(writer,
                    sheet_name=f"Due by Midnight {date.strftime('%#m-%#d-%Y')}",
                    index=False)


def generate_follow_ups(source_path, target_path):
    # work order 'Type' filter allowed column values
    work_order_type_filter = [
        "Follow-Up - Comes from a scheduled WO Checklist (PM, PdM)"
    ]

    # put the work order excel file into a dataframe from pandas
    #     and then filter out all work orders except follow ups
    #     and correctives
    df = pd.read_excel(source_path)
    df = df.loc[df["Type"].isin(work_order_type_filter)]

    # get rid of useless columns and sort the data in descending order
    #     (1 - Scheduled Start Date)
    df = df.drop(df.columns[[0, 1, 4, 5, 9, 13, 14, 15, 16]], axis=1)
    df = df.sort_values(
        by=['Sched. Start Date'],
        ascending=False)

    # create a new excel file that shows follow ups
    with pd.ExcelWriter(
        target_path,
        engine="xlsxwriter",
        engine_kwargs={'options': {'strings_to_numbers': True}}
    ) as writer:
        df.to_excel(writer,
                    sheet_name="Follow-Ups",
                    index=False)


def format(report_path, add_filters=True, delete_min_max_compliance=False):
    # use openpyxl to format the excel file
    max_compliance_report = load_workbook(report_path)
    sheet = max_compliance_report.active

    # configure alignment for the sheet
    for row in sheet[2:sheet.max_row]:
        for i in [0, 5, 6, 7]:
            cell = row[i]
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for i in [1, 2, 3, 4]:
            cell = row[i]
            cell.alignment = Alignment(
                horizontal='left', vertical='center', indent=1)

    for i, cell in enumerate(sheet[1]):
        if i in [1, 2, 3, 4]:
            cell.alignment = Alignment(
                horizontal='left', vertical='center', indent=1, wrap_text=True)
        else:
            cell.alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True)

    # set the number format of the date columns
    for row in sheet[2:sheet.max_row]:
        for i in [5, 6, 7]:
            row[i].number_format = "m/d/yyyy"

    # set the header row height
    sheet.row_dimensions[1].height = 28

    # set the background color and font size of the header
    for cell in sheet[1]:
        cell.font = Font(
            bold=True,
            color='FFFFFFFF'
        )
        cell.fill = PatternFill(
            fill_type='solid',
            fgColor='FF244062'
        )

    # set the borders of the table
    thin_border = Side(border_style='thin', color='FF000000')

    for row in sheet[1:sheet.max_row]:
        try:
            if row.value:
                row.border = Border(
                    left=thin_border,
                    right=thin_border,
                    top=thin_border,
                    bottom=thin_border
                )
        except:
            for cell in row:
                cell.border = Border(
                    left=thin_border,
                    right=thin_border,
                    top=thin_border,
                    bottom=thin_border
                )

    # if min and max compliance columns aren't needed
    if delete_min_max_compliance:
        sheet.delete_cols(7, 2)

    # autofit the column width
    for column_cells in sheet.columns:
        length = max(len(str(cell.value or "")) for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = length + 5

    # disable grid lines
    sheet.sheet_view.showGridLines = False

    # setup a filter
    if add_filters:
        sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"

    # save and close the file
    max_compliance_report.save(report_path)
    max_compliance_report.close()
