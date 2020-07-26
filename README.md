# excel-password-cracker
Simple script running through a wordlist and testing excel passwords.

Before running the script update:
1) wordlist file path (.txt file)
2) password protected excel file path
3) setting

Three different settings are available and are references by integers 1, 2 or 3
# setting 1 - runs just the wordlist
e.g.
awhile
# setting 2 - adds 1 and 2 to the end of string (2 additional password tries)
e.g.
awhile
awhile1
awhile12
# setting 3 - adds 1, 2, $ to the end of string and capitalizes first letter (6 additional tries)
e.g. 
awhile
awhile1
awhile12
awhile12$
Awhile
Awhile1
Awhile12$

Password of Test file.xlsx is "awkward"
