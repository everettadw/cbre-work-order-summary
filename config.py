from datetime import datetime


class WorkOrderSummaryConfig:
    """Config for work order summary report."""

    TARGET_DATE = datetime.today()
    TARGET_FILE_NAME = f"WO Summary {TARGET_DATE.strftime('%#m-%#d-%Y')}.xlsx"
    SOURCE_FILE_NAME = "Sheet1.xlsx"
    COLUMN_LAYOUT = {
        "Due by Midnight": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Type',
            'Sched. Start Date',
            'PM Compliance Max',
            'Reported By'
        ],
        "Scheduled for Shift": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Assigned to on Activity (Top 8 activities)',
            'PM Compliance Max'
        ],
        "Backlog": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Type',
            'Sched. Start Date',
            'PM Compliance Max',
            'Reported By'
        ],
        "Follow Ups": [
            'Work Order',
            'Equipment',
            'Description',
            '1 - WO Owner',
            'Type',
            'Reported By'
        ],
        "Training": [
            'Work Order',
            'Description',
            '1 - WO Owner',
            'Sched. Start Date'
        ]
    }
