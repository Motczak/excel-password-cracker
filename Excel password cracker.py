# -*- coding: utf-8 -*-

import win32com.client
from pythoncom import com_error
from tqdm import tqdm
import time

# Before running the script update filename with the password protected excel
# and wordlist with .txt wordlist paths
# Password of Test file.xlsx is "awkward"

filename = r"C:\Users\Hubert\Desktop\Python\Excel password\Test file.xlsx"
wordlist = r"C:\Users\Hubert\Desktop\Python\Excel password\wordlist.txt"
start = time.perf_counter()
xlApp = win32com.client.Dispatch("Excel.Application")
print("Excel library version:", xlApp.Version)

with open(wordlist, "r", encoding="utf8", errors='ignore') as file:
    for line in tqdm(file):
        password = line.rstrip()
        try:
            xlwb = xlApp.Workbooks.Open(filename, False, True, None, password)
            print()
            print("The password is: " + password)
            finish = time.perf_counter()            
            print("The script took {} seconds".format(round(finish-start, 2)))
            xlwb.Close(False)
            break
        except com_error:
            pass
