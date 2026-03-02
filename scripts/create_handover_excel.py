#!/usr/bin/env python3
"""Create Scott's Leave Handover Excel File using openpyxl"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Create workbook
wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# Define styles
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
success_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
pending_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

def style_sheet(ws, data, headers):
    """Apply styling to worksheet"""
    # Add headers
    ws.append(headers)
    
    # Style header row
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    
    # Add data rows
    for row in data:
        ws.append(row)
        for cell in ws[ws.max_row]:
            cell.alignment = left_align
            cell.border = thin_border
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width

# ============ Sheet 1: Summary ============
ws1 = wb.create_sheet("Summary")
headers1 = ['Category', 'Status', 'Owner', 'Deadline', 'Notes']
data1 = [
    ['Currency Report', 'In Progress', 'Scott', 'Ongoing', 'MC Debit pending, Visa MCC due 3 Mar'],
    ['Gaming Report Q4', 'In Progress', 'Scott', 'Wed 4 Mar', 'Pivot tables, OLS check, cost per point'],
    ['MAS 759 Reporting', 'In Progress', 'Scott', 'Next Week', 'New definition via email, UDS follow-up'],
    ['Negative Points Issue', 'Pending', 'Tech Team', 'Open', 'Follow-up sent 2 Mar'],
    ['HPE CUG Setup', 'In Progress', 'Scott/Shrini', 'Ongoing', 'Walkthrough with Grace 3 Mar 4 PM']
]
style_sheet(ws1, data1, headers1)

# ============ Sheet 2: Currency Report ============
ws2 = wb.create_sheet("Currency_Report")
headers2 = ['Data Source', 'Status', 'Last Updated', 'Notes']
data2 = [
    ['MC Credit', '✅ Complete', "Jan'26", "Dec'25 clean, Jan'26 commentary pending"],
    ['MC Debit', '⏳ Pending', '-', 'Waiting on MC for report'],
    ['Visa Credit & Debit', '⏳ Pending', '-', 'MCC data due 3 Mar 2026'],
    ['DBS/UOB/OCBC Shares', '🔄 In Progress', '-', 'Drafting with available data']
]
style_sheet(ws2, data2, headers2)

# ============ Sheet 3: Currency Actions ============
ws3 = wb.create_sheet("Currency_Actions")
headers3 = ['Action', 'Owner', 'Due Date', 'Status']
data3 = [
    ['Chase MC for Debit report', 'Scott', '3 Mar', 'Reminder set'],
    ['Visa MCC data submission', 'Visa', '3 Mar', 'Awaiting'],
    ['Draft interim report', 'Scott', '6 Mar', 'In progress'],
    ['Finalise milestones', 'Scott/Jean', '6 Mar', 'Pending Jean response']
]
style_sheet(ws3, data3, headers3)

# ============ Sheet 4: Gaming Report ============
ws4 = wb.create_sheet("Gaming_Report")
headers4 = ['Task', 'Status', 'Due Date', 'Notes']
data4 = [
    ['Pivot tables fix', '✅ Done', '27 Feb', 'Correct pulls, no dupes, sorted'],
    ['OLS extract check', '✅ Done', '27 Feb', 'Rewards verified'],
    ['Cost per point request', '✅ Done', '27 Feb', 'Requested from rewards team'],
    ['Customer letters', '✅ Done', '26 Feb', 'Prepared'],
    ['Slides', '✅ Done', '26 Feb', 'Ready for review']
]
style_sheet(ws4, data4, headers4)

# ============ Sheet 5: MAS 759 ============
ws5 = wb.create_sheet("MAS_759")
headers5 = ['Component', 'Status', 'Notes']
data5 = [
    ['MAS 759 Definition', '🔄 In Progress', 'New definition provided via email'],
    ['SA 3.1', '🔄 In Progress', 'Verification pending'],
    ['SA 5.4', '🔄 In Progress', 'Verification pending'],
    ['UDS Logic (Jiawei)', '⏳ Pending', "Follow-up needed from last week's call"]
]
style_sheet(ws5, data5, headers5)

# ============ Sheet 6: Negative Points ============
ws6 = wb.create_sheet("Negative_Points")
headers6 = ['Team', 'Contact', 'Sent Date', 'Status']
data6 = [
    ['Tech Team', 'Overall system', '2 Mar', '⏳ Awaiting reply'],
    ['Liu Wang', 'Front-end UDS', '2 Mar', '⏳ Awaiting reply'],
    ['Peter', 'Back-end capability', '2 Mar', '⏳ Awaiting reply']
]
style_sheet(ws6, data6, headers6)

# ============ Sheet 7: Deadlines During Leave ============
ws7 = wb.create_sheet("Deadlines_During_Leave")
headers7 = ['Date', 'Deadline', 'Owner', 'Action Required']
data7 = [
    ['10 Mar', 'Visa Indicative market share', 'Scott', '⚠️ Needs coverage'],
    ['11 Mar', 'MC Full Market Share', 'Scott', '⚠️ Needs coverage'],
    ['11 Mar', 'Food Rescue Day', 'Personal', '-'],
    ['14 Mar', 'Food Rescue Day (Potong Pasir)', 'Personal', '-'],
    ['15 Mar', 'Mortgage installment', 'Personal', '-'],
    ['15 Mar', 'MC Indicative market share', 'Scott', '⚠️ Needs coverage']
]
style_sheet(ws7, data7, headers7)

# ============ Sheet 8: Key Contacts ============
ws8 = wb.create_sheet("Key_Contacts")
headers8 = ['Name', 'Role', 'Topics']
data8 = [
    ['Jean', 'Manager', 'Milestones, approvals'],
    ['Shirley', 'Team Member', 'Handover (last day 4 Mar)'],
    ['Grace', 'Stakeholder', 'HPE CUG walkthrough'],
    ['Shrini', 'HPE Setup', 'HPE configuration'],
    ['Liu Wang', 'Tech Team', 'Front-end UDS'],
    ['Peter', 'Tech Team', 'Back-end auto-charging'],
    ['Sasi', 'MAS Logic', 'MAS definition clarification'],
    ['Jiawei', 'UDS', 'UDS logic follow-up']
]
style_sheet(ws8, data8, headers8)

# ============ Sheet 9: Outstanding Emails ============
ws9 = wb.create_sheet("Outstanding_Emails")
headers9 = ['Recipient', 'Topic', 'Sent Date', 'Follow-up Date']
data9 = [
    ['Ops Team', 'Gaming report address (res vs mailing)', '2 Mar', '5 Mar'],
    ['Tech Team', 'Negative point balance logic', '2 Mar', '5 Mar'],
    ['MC', 'Debit market share report', 'Pending', '3 Mar'],
    ['Visa', 'MCC data (credit & debit)', 'Pending', '3 Mar']
]
style_sheet(ws9, data9, headers9)

# ============ Sheet 10: File Locations ============
ws10 = wb.create_sheet("File_Locations")
headers10 = ['File/Location', 'Path/URL', 'Notes']
data10 = [
    ['GitHub Repo', 'https://github.com/ScottyBOS/clawd-workspace', 'Branch: clean-push'],
    ['Handover Document', 'work-deliverables/Handover_Scott_Leave_Mar2026.xlsx', 'This document'],
    ['MAS Report Tracker', 'work-deliverables/MAS_Report_Tracker_Scott_DBS.xlsx', 'Currency report tracking'],
    ['MAS Logic Mapping', 'work-deliverables/MAS_Logic_Mapping_Verification.xlsx', 'MAS definition verification'],
    ['MAS Reporting Progress', 'work-deliverables/MAS_Reporting_Progress_Tracker.xlsx', 'Progress tracker'],
    ['Gaming Reports (Shared Drive)', 'Shared/Rewards/Gaming/', 'Q4 2025 gaming reports'],
    ['Market Share Data (Shared Drive)', 'Shared/Analytics/Market_Share/', 'Visa/MC/DBS/UOB/OCBC data'],
    ['MAS Reports (Shared Drive)', 'Shared/Compliance/MAS/', 'MAS 759/SA 3.1/SA 5.4']
]
style_sheet(ws10, data10, headers10)

# ============ Sheet 11: Priority Order ============
ws11 = wb.create_sheet("Priority_Order")
headers11 = ['Priority', 'Task', 'Details']
data11 = [
    [1, 'Currency Report', 'MC Debit + Visa MCC data (time-sensitive)'],
    [2, 'MAS 759', 'New definition incorporation'],
    [3, 'Gaming Report', 'Q2 prep (early April)'],
    [4, 'Negative Points', 'Follow-up on tech responses']
]
style_sheet(ws11, data11, headers11)

# ============ Sheet 12: Notes ============
ws12 = wb.create_sheet("Notes")
headers12 = ['Topic', 'Details']
data12 = [
    ['Leave Period', '9 Mar 2026 onwards'],
    ['Expected Return', '16 Mar 2026 (Mon)'],
    ['Emergency Contact', 'Telegram'],
    ['Tokyo Trip', '8-13 Mar (limited availability)'],
    ['Escalation Path', '1. Jean (Manager) → 2. Shirley (until 4 Mar) → 3. Tech Team Lead'],
    ['Access & Permissions', 'All files on GitHub (public repo), Shared drive: Standard DBS credentials']
]
style_sheet(ws12, data12, headers12)

# Save workbook
output_file = '/home/ubuntu/clawd/work-deliverables/Handover_Scott_Leave_Mar2026.xlsx'
wb.save(output_file)

print(f"✅ Handover Excel file created: {output_file}")
print(f"\nSheets created:")
for i, sheet in enumerate(wb.sheetnames, 1):
    print(f"{i}. {sheet}")
