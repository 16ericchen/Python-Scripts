import re
# ex = 'c0002700'
# exa = re.split('([1-9\?].*)',ex,0)
# # x = re.split('(\d+)',ex)
# # print(x)
# # print(x[1])
# # print(int(x[1])+1)
# # test = ex.split('7')
# # print(test)
# print(exa)
def addOne(num):
    carry = 0
    output = ''
    for x in reversed(range(len(num))):
        if num[x].isdigit():
            checkNum = int(num[x])+carry
            if x == len(num)-1:
                checkNum = int(num[x])+carry+1
            if checkNum < 10:
                output = str(checkNum) + output
                carry = 0
            else:
                output = str(checkNum-10) + output
                carry = 1
        else:
            if carry == 1:
                output = num[x] + '1' + output
                carry = 0
            else:
                output = num[x]+output
    return output





test = [['1','2','3'],['4','5','6'],['7','8','9']]

dict = {'1':['1','2','3'],'2':['4','5','6'],'3':['7','8','9']}
# print(dict.pop('1'))
# print(dict['2'])
# print(test.remove(['1','2','3'],['4','5','6'],['7','8','9']))
string = 'CW07W00'
string2 = '67w67'
print(addOne('C909'))
print(addOne('C90W99'))
print(addOne('C799'))


# def shorten(inDict,outDict):
#         for x in inDict:
#             min = inDict[x][0][0]
#             max = inDict[x][0][0]
#             count = 0
#             newList = []
#             outDict[x] = [inDict[x][0][1],inDict[x][0][2],len(inDict[x])]
#             while count<len(inDict[x]):
#                 minNum = re.split('([1-9\?].*)',min)
#                 maxNum = re.split('([1-9\?].*)',max) 
#                 if count+1 ==len(inDict[x]):
#                     if int(maxNum[1]) - int(minNum[1]) > 1:
#                         newList.append(min+'-'+max)
#                         count+=1
#                     else:
#                             if min != max:
#                                 newList.append(min)
#                                 newList.append(max)
#                                 count +=1 
#                             else:
#                                 newList.append(min)
#                                 count+=1
#                 else:
#                     print(inDict[x])
#                     print(inDict[x][count+1][0])
#                     if ref[1].isdigit():
#                         if (ref[0]+str(int(ref[1])-1)) == max:
#                             max = inDict[x][count+1][0]
#                             count+=1
#                         else:
#                             if int(maxNum[1]) - int(minNum[1]) > 1:
#                                 newList.append(min+'-'+max)
#                                 min = inDict[x][count+1][0]
#                                 max = inDict[x][count+1][0]
#                                 count+=1
#                             else:
#                                 if min != max:
#                                     newList.append(min)
#                                     newList.append(max)
#                                     min = inDict[x][count+1][0]
#                                     max = inDict[x][count+1][0]
#                                     count +=1 
#                                 else:
#                                     newList.append(min)
#                                     min = inDict[x][count+1][0]
#                                     max = inDict[x][count+1][0]
#                                     count+=1
#                     else:
#                         ref = re.split('([1-9\?].)',ref[1])
#                         print(ref)
#             inDict[x] = newList
#         print(inDict)



test = ['1','3','4','5','7','8','10','12']
min = test[0]
max = test[0]
loopCount = 0
hyphenCount = 0
newList = []

while loopCount<len(test):
    if loopCount+1 == len(test):
        if hyphenCount > 1:
            newList.append(min+'-'+max)
            loopCount+=1
            hyphenCount = 0
        else:
            if min != max:
                newList.append(min)
                newList.append(max)
                loopCount += 1
            else:
                newList.append(min)
                loopCount += 1
    else:
        if addOne(max) == test[loopCount+1]:
            max = test[loopCount+1]
            loopCount += 1
            hyphenCount += 1
        else:
            if hyphenCount > 1:
                newList.append(min+'-'+max)
                loopCount+=1
                hyphenCount = 0
                min = test[loopCount]
                max = test[loopCount]
            else:
                if min != max:
                    newList.append(min)
                    newList.append(max)
                    hypenCount = 0
                    loopCount += 1
                    min = test[loopCount]
                    max = test[loopCount]
                else:
                    newList.append(min)
                    hyphenCount = 0
                    loopCount += 1
                    min = test[loopCount]
                    max = test[loopCount]
print(newList)