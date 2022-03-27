from datetime import date, timedelta
import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from pandas.tseries.offsets import MonthBegin, MonthEnd
import win32com.client
import inspect

MEETINGS_FLAG = 9
EMAILS_FLAG   = 6
DAY = {
    0:"MONDAY",
    1:"TUESDAY",
    2:"WEDNESDAY",
    3:"THURSDAY",
    4:"FRIDAY",
    5:"SATURDAY",
    6:"SUNDAY"
}
EXCEL_FILE_PATH = "C:/Users/sisit/Documents/טבלת צדק קשלט.xlsx"
FOLDER = "קשלט"
OUTPUT_FOLDER = "קשלט"

def get_reservations_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    #calendar = outlook.getDefaultFolder(MEETINGS_FLAG).Items   -> This is for default folder.
    calendar = outlook.getDefaultFolder(MEETINGS_FLAG).Folders("קשלט").Folders("הסתייגויות").Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')

    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def last_day_of_month(date):
    return date.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)

# For testings.
def print_attributes(a):
    for i in inspect.getmembers(a):
        if not i[0].startswith('_'):
        # Ignores methods
            if not inspect.ismethod(i[1]):
                print(i)

def print_sched(sched, days):
    for day in sched.keys():
        print(str(day), " = ", sched[day])
        if day in days["SATURDAY"]:
            print()
    print("------------------\n")
    
def get_reservations(first, last):
    cal = get_reservations_calendar(first, last)
    res = {}
    for item in cal:
        start = item.start.date()
        end = item.end.date()
        if item.AllDayEvent:
            end = end - relativedelta(days=1)
        if item.subject not in res:
            res[item.subject] = {start: end}
        else:
            res[item.subject][start] = end
    return res

def date_range(start, end):
    delta = end - start # as timedelta.
    days = [start + timedelta(days=i) for i in range(delta. days + 1)]
    return days

def get_heb_month(date_for_month):
    month = date_for_month.month
    if month == 1:
        return "ינואר"
    elif month == 2:
        return "פברואר"
    elif month == 3:
        return "מרץ"
    elif month == 4:
        return "אפריל"
    elif month == 5:
        return "מאי"
    elif month == 6:
        return "יוני"
    elif month == 7:
        return "יולי"
    elif month == 8:
        return "אוגוסט"
    elif month == 9:
        return "ספטמבר"
    elif month == 10:
        return "אוקטובר"
    elif month == 11:
        return "נובמבר"
    elif month == 12:
        return "דצמבר"
    else:
        print("Error in get_heb_month", )

def is_free(rel_date, resv):
    for start in resv.keys():
        if start <= rel_date <= resv[start]:
            return False
    return True

def get_weekends(days):
    weekends = [] 
    for i in days["THURSDAY"]:
        weekends.append([i, i + relativedelta(days=1), i + relativedelta(days=2)])
    return weekends

def assign_weekends(sched, days, rel, reservs):
    weekends = get_weekends(days) # [[date, date, date], [...], ...]
    for toran in rel.keys():
        if toran in reservs.keys():
            
            # Toran has reservations. check for free weekend.
            found = False
            for weekend in weekends:
                weekend_is_ok = True
                for day in weekend:
                    if not is_free(day, reservs[toran]):
                        weekend_is_ok = False
                        break
                if weekend_is_ok:
                    found = True
                    sched[weekend[0]] = toran
                    #sched[weekend[1]] = toran
                    #sched[weekend[2]] = toran
                    rel[toran] = True
                    weekends.remove(weekend)
                    break
            
            if not found:
                print("Error finding free weekend.")
        else:
            sched[weekends[0][0]] = toran
            #sched[weekends[0][1]] = toran
            #sched[weekends[0][2]] = toran
            weekends.remove(weekends[0])
            rel[toran] = True

def get_relevant(date_for_month):
    # Get for month all weekend and weekdays closing people from excel.
    xf = pd.read_excel(EXCEL_FILE_PATH).to_dict()
    weekends = {}
    weekdays = {}

    heb_month = get_heb_month(date_for_month)
    for i in range(0, len(xf["שם מלא"])):
        name = list(xf["שם מלא"].values())[i]
        toran = list(xf[heb_month].values())[i]
        if toran == "לילה":
            weekdays[name] = False
        elif toran == "שבת":
            weekends[name] = False
    return {"weekends": weekends, "weekdays": weekdays}

def assign_toran(sched, days, toran, reservs, is_ok):
    if is_ok:
        # Search from sunday to wednesday
        for weekday in ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY"]:
            for day in days[weekday]:
                if day in sched and not sched[day]:
                    if toran in reservs:
                        if is_free(day, reservs[toran]):
                            sched[day] = toran
                            return True
                    else:
                        sched[day] = toran
                        return True

    else:
        # Search from wednesday to sunday.
        for weekday in ["WEDNESDAY", "TUESDAY", "MONDAY", "SUNDAY"]:
            for day in days[weekday]:
                if day in sched and not sched[day]:
                    if toran in reservs:
                        if is_free(day, reservs[toran]):
                            sched[day] = toran
                            return True
                    else:
                        sched[day] = toran
                        return True

    print("Error finding weekday for ", toran)
    return False

def get_outlook_assigned(folder, begin, end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    #calendar = outlook.getDefaultFolder(MEETINGS_FLAG).Items   -> This is for default folder.
    calendar = outlook.getDefaultFolder(MEETINGS_FLAG).Folders(folder).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')

    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    assigned = {}
    for item in calendar:
        start = item.start.date()
        end = item.end.date()
        if item.AllDayEvent:
            end = end - relativedelta(days=1)
        if item.subject not in assigned:
            assigned[item.subject] = {start: end}
        else:
            assigned[item.subject][start] = end
    return assigned

def get_non_assigned_relevant(assigned, relevant):
    non_assigned = {"weekdays": relevant["weekdays"].copy(), "weekends": relevant["weekends"].copy()}
    torans = relevant["weekdays"].keys()
    for toran in torans:
        if toran in assigned:
            non_assigned["weekdays"].pop(toran)

    torans = relevant["weekends"].keys()
    for toran in torans:
        if toran in assigned:
            non_assigned["weekends"].pop(toran)
    
    return non_assigned

def get_non_assigned_sched(outlook_assigned, sched):
    non_assigned = sched.copy()

    for toran in outlook_assigned.keys():
        date = list(outlook_assigned[toran].keys())[0]
        if date in sched.keys():
            non_assigned.pop(date)

    return non_assigned

if __name__ == '__main__':
    today = date.today()
    first = date.today().replace(day=1) + relativedelta(months=1) 
    last = last_day_of_month(first) 
    fix_last = last + relativedelta(days=1)
    fix_first = first - relativedelta(days=1)
    
    # Get Reservations.
    reservs = get_reservations(fix_first, fix_last)
    
    # Get relevant from excel without those already assigned in outlook.
    relevant = get_relevant(first) 
    outlook_assigned = get_outlook_assigned(FOLDER, fix_first, fix_last)
    relevant = get_non_assigned_relevant(outlook_assigned.keys(), relevant)
    # relevant = {"weekends": {"name": False}, "weekdays": {"name": False}}

    # Create month schedule.
    sched = {}
    days = {}
    for i in date_range(first, last):
        sched[i] = ""
        if DAY[i.weekday()] not in days.keys():
            # TODO: if not in outlook__assigned.
            days[DAY[i.weekday()]] = [i]
        else:
            # TODO: if not in outlook__assigned.
            days[DAY[i.weekday()]].append(i)

    print(sched)
    sched = get_non_assigned_sched(outlook_assigned, sched)
    print("-------------")
    print(sched)

    assign_weekends(sched, days, relevant["weekends"], reservs)
    
    blacklist = open("blacklist.txt", "r", encoding="utf-8").read().splitlines()
    whitelist = open("whitelist.txt", "r", encoding="utf-8").read().splitlines()

    # Enter whitelist.
    for toran in whitelist:
        if toran in relevant["weekdays"]:
            relevant["weekdays"][toran] = assign_toran(sched, days, toran, reservs, True)

    # Enter reservs if relevant.
    for toran in relevant["weekdays"]:
        if relevant["weekdays"][toran] == False and toran in reservs.keys():
            if toran in blacklist:
                relevant["weekdays"][toran] = assign_toran(sched, days, toran, reservs, False)
            else:
                relevant["weekdays"][toran] = assign_toran(sched, days, toran, reservs, True)

    # Enter blacklist.
    for toran in blacklist:
        if toran in relevant["weekdays"] and relevant["weekdays"][toran] == False:
            relevant["weekdays"][toran] = assign_toran(sched, days, toran, reservs, False)

     # Enter all the rest.
    for toran in relevant["weekdays"]:
        if relevant["weekdays"][toran] == False:
            relevant["weekdays"][toran] = assign_toran(sched, days, toran, reservs, True)

    print_sched(sched, days)
    
    # Add items to outlook.
    outlook = win32com.client.Dispatch("Outlook.Application")
    calendar = outlook.GetNamespace('MAPI').getDefaultFolder(MEETINGS_FLAG).Folders(OUTPUT_FOLDER)   #-> This is for default folder.

    for day in sched.keys():
        if sched[day]:
            appt = calendar.Items.Add(1) # AppointmentItem
            appt.Start = str(day) # "yyyy-mm-dd"
            appt.Subject = sched[day]
            appt.AllDayEvent = True
            appt.Organizer = ""
            if DAY[day.weekday()] == "THURSDAY":
                appt_days = 3
            else:
                appt_days = 1
            appt.Duration = 1440 * appt_days # In minutes * days
            appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
            
            #appt.Recipients.Add("test@test.com") # Don't end ; as delimiter
            appt.Save()
            #appt.Send()
