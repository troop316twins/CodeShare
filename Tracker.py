"""
Track When a file was modified.  To be run hourly?
"""
import os.path
import time

FileA = r"N:\Projects\029-030 Aloha CV3600 Containership\300 Engineering & Technical\10_Plan Approval\029 ALOHA CV3600 Plan Approval Tracking.xlsx"
OutputFile = ""
#input the output file
# if date is different from last date in the output file:
# append the new date to that output file
# else do nothing
print (time.ctime(os.path.getmtime(FileA)))
