import os
import find

#run in terminal with Python med_aid_script.py

#choose file
file_path = input("Enter file path: ")
#find info:
content=find.iterate(file_path)
print('done')
print('content: '+content)
find.format(content)
print("done")