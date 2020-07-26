# -*- coding: utf-8 -*-

import win32com.client
from pythoncom import com_error
import time

# Before running the script update filename with the password protected excel
# and wordlist with .txt wordlist paths
# Password of Test file.xlsx is "awkward"
# You can select between 3 different settings by
# changing setting variable to be integer 1, 2 or 3
# setting 1 - runs just the wordlist
# setting 2 - adds 1 and 2 to the end of string
# setting 3 - setting 2 + capitalizes first letter and adds "$" at the end

# ------------------------- Settings -------------------------------------
filename = r"C:\Users\Hubert\Desktop\Python\Excel password\Test file.xlsx"
wordlist = r"C:\Users\Hubert\Desktop\Python\Excel password\wordlist.txt"
setting = 1
# ------------------------------------------------------------------------

# Nothing beyond this line needs to be manually adjusted
start = time.perf_counter()
counter = 0
found_password = False
xlApp = win32com.client.Dispatch("Excel.Application")
print("Excel library version:", xlApp.Version)


def try_password(password):
    global counter
    global found_password
    counter += 1
    try:
        xlwb = xlApp.Workbooks.Open(filename, False, True, None, password)
        print()
        print("The password is: " + password)
        finish = time.perf_counter()
        print("The script took {} seconds".format(round(finish-start, 2)))
        print("{} passwords were tested.".format(counter))
        xlwb.Close(False)
        found_password = True
    except com_error:
        print(password)


# setting 1 - runs just the wordlist
# setting 2 - computes different permutations adds 1 and 2 to the end of string
# setting 3 - setting 2 + capitalizes first letter and adds "$" at the end

def get_password(line, setting):
    global found_password
    if setting == 1:
        passwords = [line.rstrip()]
        return passwords
    elif setting == 2:
        passwords = [line.rstrip()]
        passwords.append(line.rstrip() + str(1))
        passwords.append(line.rstrip() + str(1) + str(2))
        return passwords
    elif setting == 3:
        passwords = [line.rstrip()]
        passwords.append(line.rstrip() + str(1))
        passwords.append(line.rstrip() + str(1) + str(2))
        passwords.append(line.rstrip() + str(1) + str(2) + "$")
        passwords.append(line.rstrip().capitalize())
        passwords.append(line.rstrip().capitalize() + str(1))
        passwords.append(line.rstrip().capitalize() + str(1) + str(2) + "$")
        return passwords
    else:
        print("Non existing setting selected - please choose 1, 2 or 3.")
        found_password = True


def main(setting):
    with open(wordlist, "r", encoding="utf8", errors='ignore') as file:
        for line in file:
            passwords = get_password(line, setting)
            if found_password is True:
                break
            for password in passwords:
                try_password(password)
                if found_password is True:
                    break


if __name__ == '__main__':
    main(setting)
