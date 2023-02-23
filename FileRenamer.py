import os,re
from tkinter.filedialog import askdirectory
path = askdirectory(title='Select Folder') # shows dialog box and return the path
# for i in range(10):
#     f = open("myfile"+"["+str(i)+"]"+".txt", "x")
for file in os.listdir(path):
    if "(" or "[" in file:
        t = re.sub("[\(\[].*?[\)\]]", "", file)
        if not os.path.exists(t):
            os.rename(file,str()+t)


