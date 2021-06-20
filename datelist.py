import master
import datetime
import calendar

# get the current month
curt_month = datetime.datetime.now().month
# print(curt_month)
# get the current year to find the leap year it's necessary
curt_year = datetime.datetime.now().year
# print(curt_year)

# condition to get the dec month separately as 1-1 will be zero and for other month days
if curt_month == 1:
    numdays = calendar.monthrange(year=curt_year,month=12)[1]
else:
    numdays = calendar.monthrange(year=curt_year,month=curt_month-1)[1]
# print(numdays)
cur_numdays = calendar.monthrange(year=curt_year, month=curt_month)[1]

# store the dates from constant n = 16th of previous month to 15th of current as per payroll cycle
datelist = []
master.n = 16

# change the current date to 16th
base = datetime.datetime.today().replace(day=master.n)

for i in range(1,numdays+1):
    # its reducing one day from the base date 16th and storing in datelist from 15th of the current month
    tempdate = base - datetime.timedelta(days=i)
    datelist.append(tempdate.strftime("%d/%m/%Y"))
    master.n -= 1
# ['15/05/2021', '14/05/2021', '13/05/2021', '12/05/2021', '11/05/2021', '10/05/2021', '09/05/2021', '08/05/2021', '07/05/2021', '06/05/2021', '05/05/2021', '04/05/2021', '03/05/2021', '02/05/2021', '01/05/2021', '30/04/2021', '29/04/2021', '28/04/2021', '27/04/2021', '26/04/2021', '25/04/2021', '24/04/2021', '23/04/2021', '22/04/2021', '21/04/2021', '20/04/2021', '19/04/2021', '18/04/2021', '17/04/2021', '16/04/2021']

# will reverse the dates so last month 16 till number of days will be final list
datelist.reverse()
# print(datelist)
# ['16/04/2021', '17/04/2021', '18/04/2021', '19/04/2021', '20/04/2021', '21/04/2021', '22/04/2021', '23/04/2021', '24/04/2021', '25/04/2021', '26/04/2021', '27/04/2021', '28/04/2021', '29/04/2021', '30/04/2021', '01/05/2021', '02/05/2021', '03/05/2021', '04/05/2021', '05/05/2021', '06/05/2021', '07/05/2021', '08/05/2021', '09/05/2021', '10/05/2021', '11/05/2021', '12/05/2021', '13/05/2021', '14/05/2021', '15/05/2021']