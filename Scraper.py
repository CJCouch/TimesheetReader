import win32com.client
import re

# Returns list of values based on search regex
def scrape_msg(msg, regex):
    body = msg.Body
    print("\n{0}\nReading timesheet from {1}\n{0}\n{2}".format(
        '='*100, str(msg), body))
    # Delete all the non times from list
    times = re.findall(regex, body)
    return times

def open_email(path):
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    return outlook.OpenSharedItem(path)

# Formats list of times to meet spreadsheet requirements
def format_times(times):
    days = times
    for i in range(0, len(days)):
        if 'p.m.' in str(days[i]):
            s = str(days[i]).replace(' p.m.', '')
            days[i] = s

        if ':' not in str(days[i]):
            s = str(days[i]) + ':00'
            days[i] = s

        if '-' in str(days[i]):
            s = str(days[i]).replace('-', '0')
            days[i] = s
    return days
