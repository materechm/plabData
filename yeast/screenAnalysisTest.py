import xlsxwriter
import csv
import numpy
from pyper import *
from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf
import math
from collections import defaultdict

def get_data(filePath, fileName):
  rb = open_workbook(filePath)
  r_sheet = rb.sheet_by_index(0)
  write_file = open(fileName, 'w')
  for row_index in range(0, r_sheet.nrows):
    if r_sheet.cell(row_index,0).value == "Plate:":
        plateName = r_sheet.cell(row_index,1).value
        row = row_index+2
        nControl = []
        pControl = []
        sig1 = []
        sig2 = []
        for i in range(1, 17):
            nControl.append(r_sheet.cell(row,2).value)
            nControl.append(r_sheet.cell(row,3).value)
            pControl.append(r_sheet.cell(row,24).value)
            pControl.append(r_sheet.cell(row,25).value)
            nControlAve = numpy.mean(nControl)
            nControlStdev = numpy.std(nControl)
            pControlAve = numpy.mean(pControl)
            pControlStdev = numpy.std(pControl)
            upper = nControlAve + nControlStdev*3
            lower = pControlAve - pControlStdev*3
            for j in range (3,23):
                if r_sheet.cell(row,j+1).value > upper:
                    sig1.append(str(i)+";"+str(j))
                    if r_sheet.cell(row,j+1).value > lower:
                        sig2.append(str(i)+";"+str(j))
        write_file.write("plate,"+plateName+"\n")
        write_file.write("sig1,")
        for x in sig1:
            write_file.write('%s,' %x)
        write_file.write("\n")
        write_file.write("sig2,")
        for x in sig2:
            write_file.write('%s,' %x)
        write_file.write("\n")
        row += 1

path23h = "/Users/mtchavez/plabData/yeast/1stScreenData/23h.xlsx"
path48h = "/Users/mtchavez/plabData/yeast/1stScreenData/48h.xlsx"
fileName23h = "/Users/mtchavez/plabData/yeast/1stScreenData/data23h.csv"
fileName48h = "/Users/mtchavez/plabData/yeast/1stScreenData/data48h.csv"
get_data(path23h, fileName23h)
get_data(path48h, fileName48h)

def compare_data(fileName, finalFile):
    read_file = open(fileName)
    data = {}
    plateName = ""
    reader = csv.reader(read_file, delimiter=',')
    for row in reader:
        if row[0] == "plate":
            plateName = row[1]
        if row[0] == "sig1":
            if len(row) > 1:
                sig1 = [x for x in row[1:]]
            else:
                sig1 = []
        if row[0] == "sig2":
            if len(row) > 1:
                sig2 = [x for x in row[1:]]
            else:
                sig2 = []
            data[plateName] = sig1, sig2
    read_file.close()
    write_file = open(finalFile, 'w')
    A = {}
    B = {}
    C = {}
    D = {}
    E = {}
    F = {}
    plates = set()
    for key in data:
        plates.add(key[:-5])
    for plate in plates:
        plate = [(x, data[x]) for x in data if x.startswith(plate)]
        for item in plate:
            subplate = item[0]
            subplateData = item[1]
            subplateSig1 = subplateData[0]
            subplateSig2 = subplateData[1]
            if subplate[-5] == "A":
                A[subplate[:-5]] = subplateSig1, subplateSig2
            if subplate[-5] == "D":
                D[subplate[:-5]] = subplateSig1, subplateSig2
            if subplate[-5] == "B":
                B[subplate[:-5]] = subplateSig1, subplateSig2
            if subplate[-5] == "E":
                E[subplate[:-5]] = subplateSig1, subplateSig2
            if subplate[-5] == "C":
                C[subplate[:-5]] = subplateSig1, subplateSig2
            if subplate[-5] == "F":
                F[subplate[:-5]] = subplateSig1, subplateSig2
    tdp43wt = defaultdict(list)
    tdp43m337v = defaultdict(list)
    fuswt = defaultdict(list)
    ts1 = {}
    ts2 = {}
    tms1 = {}
    tms2 = {}
    fs1 = {}
    fs2 = {}
    for k, v in A.items() + D.items():
        tdp43wt[k].append(v)
    for k, v in B.items() + E.items():
        tdp43m337v[k].append(v)
    for k, v in C.items() + F.items():
        fuswt[k].append(v)
    for key in tdp43wt:
        try:
            sig11 = tdp43wt[key][0][0]
            sig12 = tdp43wt[key][1][0]
            sig21 = tdp43wt[key][0][1]
            sig22 = tdp43wt[key][1][1]
            sig1 = [x for x in sig11 if x in sig12]
            sig2 = [x for x in sig21 if x in sig22]
            ts1[key] = sig1
            ts2[key] = sig2
            break
        except (IndexError):
            print tdp43wt[key]
    for key in tdp43m337v:
        try:
            sig11 = tdp43m337v[key][0][0]
            sig12 = tdp43m337v[key][1][0]
            sig21 = tdp43m337v[key][0][1]
            sig22 = tdp43m337v[key][1][1]
            sig1 = [x for x in sig11 if x in sig12]
            sig2 = [x for x in sig21 if x in sig22]
            tms1[key] = sig1
            tms2[key] = sig2
            break
        except (IndexError):
            print tdp43m337v[key]
    for key in fuswt:
        try:
            sig11 = fuswt[key][0][0]
            sig12 = fuswt[key][1][0]
            sig21 = fuswt[key][0][1]
            sig22 = fuswt[key][1][1]
            sig1 = [x for x in sig11 if x in sig12]
            sig2 = [x for x in sig21 if x in sig22]
            fs1[key] = sig1
            fs2[key] = sig2
        except (IndexError):
            print fuswt[key]
    write_file.write("TDP43WT"+"\n")
    write_file.write("sig1"+"\n")
    for key in ts1:
        write_file.write(key + ",")
        for x in ts1[key]:
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.write("sig2"+"\n")
    for key in ts2:
        for x in ts2[key]:
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.write("TDP43M337V"+"\n")
    write_file.write("sig1"+"\n")
    for key in tms1:
        write_file.write(key + ",")
        for x in tms1[key]:
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.write("sig2"+"\n")
    for key in tms2:
        write_file.write(key + ",")
        for x in tms2[key]:
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.write("FUSWT"+"\n")
    write_file.write("sig1"+"\n")
    for key in fs1:
        write_file.write(key + ",")
        for x in fs1[key]:
            print x + "fs1[key]"
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.write("sig2"+"\n")
    for key in fs2:
        write_file.write(key + ",")
        for x in fs2[key]:
            print x + "fs2[key]"
            write_file.write('%s,' %x)
        write_file.write("\n")
    write_file.close()

finalFile23h = "/Users/mtchavez/plabData/yeast/1stScreenData/results23h.csv"
finalFile48h = "/Users/mtchavez/plabData/yeast/1stScreenData/results48h.csv"
compare_data(fileName23h, finalFile23h)
compare_data(fileName48h, finalFile48h)
