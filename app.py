from pathlib import Path
import openpyxl as xl
import os
import master
import datelist


print("Payroll Automation.")
print("No of days " + str(datelist.numdays))
print("Total 6 Steps")

import stepone
print(stepone.status)

import steptwo
print(steptwo.status)

import stepthree
print(stepthree.status)

import stepfour
print(stepfour.status)

import stepfive
print(stepfive.status)

import stepsix
print(stepsix.status)

import stepseven
print(stepseven.status)

import stepeight
print(stepeight.status)

print("payroll and exception files has been created")