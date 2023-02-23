from pynput import keyboard
import time
import sqlite3,datetime
from sqlite3 import Error
from consolemenu import *
from consolemenu.items import *
from openpyxl import Workbook
# im trying to assign the variables values
time = datetime.date.today()
projectDate = time.strftime("%m-%d-%Y")
totalList = [] 
projectName=projectAssembly=serialNumber=projectMatrix=projectNote='what'
testList = [projectName,projectAssembly,projectMatrix,projectNote,serialNumber,projectDate]
inputList = ["Input Project Name:","Input Assembly:","Input Matrix","Input Note","Scan Serial Number:"]
def assignInput(i):
    print(inputList[i])
    testList[i]= input()
    print('assign',testList[i])
    return 
def enterData():
        global serialNumber,projectAssembly,projectMatrix,projectName,projectNote
        while 1:
            counter = 0
            x = True
            for y in range(len(inputList)):
                assignInput(y)
            serialNumber = testList[4]
            totalList.append(serialNumber)
            while x is True:
                counter += 1
                print("Current Number of Boards: "+str(counter))
                print("Scan Again, Press E to Exit, Press N to Input New Project, Press Q to change Note, Press Z to change Matrix,Press D to delete Previous Entry")
                serialNumber = input()
                if serialNumber == 'n':
                    x = False
                elif serialNumber == 'e':
                    print(totalList)
                    # return 
                elif serialNumber == 'q':
                    assignInput(3)
                    assignInput(4)
                    serialNumber = testList[4]
                    totalList.append(serialNumber)
                elif serialNumber == 'd':
                    totalList.pop()
                    counter -= 2
                elif serialNumber == 'z':
                    assignInput(2)
                    assignInput(4)
                    serialNumber = testList[4]
                    totalList.append(serialNumber)
                else:
                    totalList.append(serialNumber)
            
def main():
        menu = ConsoleMenu("Title","Subtitle")
        function_Data = FunctionItem("Enter New Project Data",enterData)
        menu.append_item(function_Data)
        menu.show()

if __name__ == '__main__':
    main()



# def on_press(key):
#     print (key)
#     if key == keyboard.Key.end:
#         print ('end pressed')
#         return False        
# with keyboard.Listener(on_press=on_press) as listener:
#     while True:
#         print ('program running')
#         time.sleep(5)
#     listener.join()

# break_program = False
# def on_press(key):
#     global break_program
#     print (key)
#     if key == keyboard.Key.delete:
#         print ('end pressed')
#         break_program = True
#         return False

# with keyboard.Listener(on_press=on_press) as listener:
#     while break_program == False:
#         print ('program running')
#     listener.join()