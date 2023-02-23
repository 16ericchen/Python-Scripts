from os.path import exists,splitext
import os

x,y,spi =[],[],[]
file_list = ["SPI-TOP.txt","SPI-BOT.txt"]
omit = ["PP","XW","TP"]

def writetotextfile(side,tf):
        for d in side:
            tf.write(str(d)+"\t")
        tf.write('\n')   

def topandbot(sorted,topfile,botfile):
    for side in sorted:
        if "NO" in side[-1]:
            writetotextfile(side[:-1],topfile)
        if "YES" in side[-1]:
            writetotextfile(side[:-1],botfile)
    topfile.close(),botfile.close()

def createSPI(inputArray):
    for line in inputArray:
        x.append(line.strip().split(','))
    for item in x[5:]:
        input = item[0:1]+item[-4:-1]+['T']
        if item[0][:2] not in omit:
            if (item[1][0:7] != "SMT_PAD") and ("SKT" not in item[-5]):
                spi.append(input+item[-5:-4]+item[-1:])
        if item[1][0:7]=="TP_5016":
            spi.append(input+item[-5:-4]+item[-1:])
    inputArray.close()
    x.clear()


def createFiles(inputFile,outputFile1,outputFile2):
    Top = open(outputFile1,'w')
    Bot = open(outputFile2,'w')
    f = open(inputFile,"r",encoding='utf-8')
    createSPI(f)
    topandbot(spi,Top,Bot)

def changExtension():
    for file in file_list:
        if exists(file):
            base = splitext(file)[0]
            if not exists(base+'.aoi'):
                os.rename(file,base+'.aoi')

if exists("comp.txt"):
    createFiles("comp.txt","SPI-TOP.txt","SPI-BOT.txt")
    changExtension()