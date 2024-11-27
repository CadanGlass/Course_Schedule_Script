import win32com.client
import re
from datetime import datetime, timedelta, date
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from collections import defaultdict
import win32timezone

def get_emails(from_date, subject_keyword):
    """
    Fetches emails from Outlook Inbox sent on or after `from_date`
    containing `subject_keyword` in their subject line.
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[SentOn]", Descending=False)
        restrict_date = from_date.strftime('%m/%d/%Y 12:00 AM')
        restricted = messages.Restrict(f"[SentOn] >= '{restrict_date}'")
        filtered_emails = [email for email in restricted if subject_keyword.lower() in email.Subject.lower()]
        return filtered_emails
    except Exception as e:
        print(f"Error fetching emails: {e}")
        return []

def parse_email(email_body):
    """
    Parses email body to extract original sent date and counts of changes.
    """
    sent_date_match = re.search(r"Sent:\s+(.*)", email_body, re.IGNORECASE)
    sent_datetime = None
    if sent_date_match:
        sent_date_str = sent_date_match.group(1).strip()
        date_formats = [
            '%A, %B %d, %Y %I:%M %p', '%B %d, %Y %I:%M %p',
            '%a, %b %d, %Y %I:%M %p', '%d %B %Y %I:%M %p'
        ]
        for fmt in date_formats:
            try:
                sent_datetime = datetime.strptime(sent_date_str, fmt)
                break
            except ValueError:
                continue
    sent_date = sent_datetime.date() if sent_datetime else None
    additions_match = re.search(r"Additions:\s*(.*?)\s*(Cancellations:|Instructor Changes:|$)", email_body, re.IGNORECASE | re.DOTALL)
    cancellations_match = re.search(r"Cancellations:\s*(.*?)\s*(Instructor Changes:|$)", email_body, re.IGNORECASE | re.DOTALL)
    instructor_changes_match = re.search(r"Instructor Changes:\s*(.*)", email_body, re.IGNORECASE | re.DOTALL)

    def count_valid_lines(section_text, pattern=None):
        if not section_text:
            return 0
        lines = section_text.strip().split("\n")
        return len([line for line in lines if line.strip() and (not pattern or re.match(pattern, line.strip()))])

    additions = count_valid_lines(additions_match.group(1) if additions_match else None, r".+\s*/\s*CRN\s*\d+")
    cancellations = count_valid_lines(cancellations_match.group(1) if cancellations_match else None, r".+\s*/\s*CRN\s*\d+")
    instructor_changes = count_valid_lines(instructor_changes_match.group(1) if instructor_changes_match else None, r".+\s*/\s*CRN\s*\d+\s*-\s*.+")
    return sent_date, additions, cancellations, instructor_changes

def write_to_excel(data, weekly_data, total_counts, averages, email_details, output_file):
    """
    Writes aggregated data to an Excel file with detailed statistics.
    """
    wb = openpyxl.Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    center_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    ws_daily = wb.active
    ws_daily.title = "Daily Changes Summary"
    headers = ["Date", "Section Additions", "Section Cancellations", "Instructor Changes", "Note"]
    ws_daily.append(headers)
    for col in range(1, 6):
        cell = ws_daily.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border

    sorted_dates = sorted(data.keys())
    for date_obj in sorted_dates:
        ws_daily.append([
            date_obj, data[date_obj]['additions'], data[date_obj]['cancellations'],
            data[date_obj]['instructor_changes'],
            "Multiple emails processed" if data[date_obj]['email_count'] > 1 else ""
        ])
    for row_cells in ws_daily.iter_rows(min_row=2, max_row=ws_daily.max_row, min_col=1, max_col=5):
        for cell in row_cells:
            cell.alignment = center_alignment
            cell.border = thin_border

    ws_weekly = wb.create_sheet(title="Weekly Changes Summary")
    headers = ["Week Start Date", "Week End Date", "Date Range", "Section Additions", "Section Cancellations", "Instructor Changes"]
    ws_weekly.append(headers)
    for col in range(1, 7):
        cell = ws_weekly.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border

    sorted_weeks = sorted(weekly_data.keys())
    for week_start_date in sorted_weeks:
        week_end_date = week_start_date + timedelta(days=6)
        ws_weekly.append([
            week_start_date, week_end_date,
            f"{week_start_date.strftime('%b %d, %Y')} - {week_end_date.strftime('%b %d, %Y')}",
            weekly_data[week_start_date]['additions'],
            weekly_data[week_start_date]['cancellations'],
            weekly_data[week_start_date]['instructor_changes']
        ])
    for row_cells in ws_weekly.iter_rows(min_row=2, max_row=ws_weekly.max_row, min_col=1, max_col=6):
        for cell in row_cells:
            cell.alignment = center_alignment
            cell.border = thin_border

    ws_stats = wb.create_sheet(title="Overall Statistics")
    ws_stats.cell(row=1, column=1, value="Total Counts").font = header_font
    row = 2
    for key, value in total_counts.items():
        ws_stats.cell(row=row, column=1, value=key).font = Font(bold=True)
        ws_stats.cell(row=row, column=2, value=value)
        row += 1
    ws_stats.cell(row=row, column=1, value="Averages per Day").font = header_font
    row += 1
    for key, value in averages.items():
        ws_stats.cell(row=row, column=1, value=key).font = Font(bold=True)
        ws_stats.cell(row=row, column=2, value=round(value, 2))
        row += 1

    ws_emails = wb.create_sheet(title="Email Details")
    headers = ["Email Sent On", "Subject", "Parsed Sent Date", "Section Additions", "Section Cancellations", "Instructor Changes"]
    ws_emails.append(headers)
    for col in range(1, 7):
        cell = ws_emails.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border

    for email in email_details:
        ws_emails.append([
            email['email_sent_on'], email['subject'], email['sent_date'],
            email['additions'], email['cancellations'], email['instructor_changes']
        ])
    for row_cells in ws_emails.iter_rows(min_row=2, max_row=ws_emails.max_row, min_col=1, max_col=6):
        for cell in row_cells:
            cell.alignment = center_alignment
            cell.border = thin_border

    try:
        wb.save(output_file)
        print(f"Data successfully written to {output_file}")
    except Exception as e:
        print(f"Failed to save the workbook: {e}")

if __name__ == "__main__":
    from_date = datetime(2024, 11, 15)
    subject_keyword = "Changes to the 202510 Course Schedule"
    output_file = "Course_Changes_Summary.xlsx"

    emails = get_emails(from_date, subject_keyword)
    if not emails:
        print("No matching emails found.")
    else:
        data = defaultdict(lambda: {'additions': 0, 'cancellations': 0, 'instructor_changes': 0, 'email_count': 0})
        email_details = []

        for email in emails:
            try:
                sent_date, additions, cancellations, instructor_changes = parse_email(email.Body)
                if sent_date is None:
                    sent_date = email.SentOn.date()
                data[sent_date]['additions'] += additions
                data[sent_date]['cancellations'] += cancellations
                data[sent_date]['instructor_changes'] += instructor_changes
                data[sent_date]['email_count'] += 1
                email_details.append({
                    'sent_date': sent_date,
                    'email_sent_on': email.SentOn.replace(tzinfo=None) if email.SentOn.tzinfo else email.SentOn,
                    'subject': email.Subject,
                    'additions': additions,
                    'cancellations': cancellations,
                    'instructor_changes': instructor_changes,
                })
            except Exception as e:
                print(f"Error processing an email: {e}")

        if data:
            total_counts = {
                'Total Additions': sum(entry['additions'] for entry in data.values()),
                'Total Cancellations': sum(entry['cancellations'] for entry in data.values()),
                'Total Instructor Changes': sum(entry['instructor_changes'] for entry in data.values())
            }
            days_count = len(data)
            averages = {
                'Average Additions per Day': total_counts['Total Additions'] / days_count if days_count else 0,
                'Average Cancellations per Day': total_counts['Total Cancellations'] / days_count if days_count else 0,
                'Average Instructor Changes per Day': total_counts['Total Instructor Changes'] / days_count if days_count else 0
            }
            weekly_data = defaultdict(lambda: {'additions': 0, 'cancellations': 0, 'instructor_changes': 0})
            for date_obj, counts in data.items():
                week_start_date = date_obj - timedelta(days=date_obj.weekday())
                weekly_data[week_start_date]['additions'] += counts['additions']
                weekly_data[week_start_date]['cancellations'] += counts['cancellations']
                weekly_data[week_start_date]['instructor_changes'] += counts['instructor_changes']
            write_to_excel(data, dict(weekly_data), total_counts, averages, email_details, output_file)
        else:
            print("No data to write after processing emails.")
