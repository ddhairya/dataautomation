from pathlib import Path
import datetime
# start date for payroll
n = 16
# work hours for making exception
max_work_hr_exception = 11.0
normal_work_hr = 9.0
normal_ot_hr = 2.0
normal_ot_excp_hr = 0.5


# month
cur_month = datetime.datetime.now().strftime("%b")
cur_year = datetime.datetime.now().strftime("%Y")
# company name for running the automation
company = "BRIOCH"
company_code = "40"

# employee code len
emp = 5
emp_start_code = "4"

file = Path("c:/it/PayrollAuto/empshiftsum.xlsx")
filepath = "c:\\it\\PayrollAuto\\"

finalpath = "c:/it/PayrollAuto/"

temppath = "c:/it/PayrollAuto/working_tmp/"
tempfile = "c:\\it\\PayrollAuto\\working_tmp\\"

masterpath = "c:/it/PayrollAuto/master/"
masterfile = "c:\\it\\PayrollAuto\\master\\"