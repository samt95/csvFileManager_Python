#!/opt/bin/python3

#############################################
#
# Python CSV File Manager
#
# Developer: Sam Taylor
# 
#############################################

import csv
import sys
import re

class spreadsheet:

	def __init__(self, fn, nr, nc, con):
		self.fileName = fn
		self.numRows = nr
		self.numCols = nc
		self.contents = con

#global spreadsheet
ss = spreadsheet('', 0, 0, [])


########################################


def load(command):
	if len(command) == 1:
		print("-- filename not specified")
	else:
		command[1] = hasQuotes(command[1])
		ss_load(command[1])


def printrow(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- row not specified")
		return
	row1 = getColRowArg(command[1])
	if len(command) == 2:
		ss_printrow(row1) #print 1 row
	else:
		row2 = getColRowArg(command[2])
		ss_printrows(row1, row2) #print multiple rows


def evalSum(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- column not specified")
		return
	col = getColRowArg(command[1])
	if col<0 or col>=ss.numCols:
		print("-- column (",command[1],") is out of range")
		return
	print("-- sum =", ss_evalSum(col))


def evalAvg(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- column not specified")
		return
	col = getColRowArg(command[1])
	if col<0 or col>=ss.numCols:
		print("-- column (",command[1],") is out of range")
		return
	print("-- average =", ss_evalAvg(col))


def findRow(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- column and string not specified")
	elif len(command) == 2:
		print("-- column or string not specified")
	elif len(command) == 3:
		col = getColRowArg(command[1])
		command[2] = hasQuotes(command[2])
		try:
			ss_findRow(col, command[2], 0)
		except:
			print("-- bad column specification")
	else:
		col = getColRowArg(command[1])
		command[2] = hasQuotes(command[2])
		startRow = getColRowArg(command[3])
		try:
			ss_findRow(col, command[2], startRow)
		except:
			print("-- bad column or row specification")


def sortNumeric(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- column not specified")
	else:
		col = getColRowArg(command[1])
		ss_sortNumeric(col)		


def sort(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- column not specified")
	else:
		col = getColRowArg(command[1])
		ss_sort(col)


def save(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		ss_save(ss.fileName)
	else:
		command[1] = hasQuotes(command[1])
		ss_save(command[1])


def merge(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- merge file not specified")
	else:
		command[1] = hasQuotes(command[1])
		ss_merge(command[1])


def deleteRow(command):
	if isFileLoaded() == False:
		return
	if len(command) == 1:
		print("-- row not specified")
		return
	row = getColRowArg(command[1])
	if row < 0 or row >= ss.numRows:
		print("-- row (", row, ") is out of range")
	else:
		ss_deleteRow(row)


def unload():
	ss.fileName = ''
	ss.numRows = 0
	ss.numCols = 0
	ss.contents = []


def printStats():
	print("File:", ss.fileName)
	print("Rows:", ss.numRows)
	print("Columns:", ss.numCols)	


##############################################


def ss_load(fileName):
	try:
		csvfile = open(fileName, newline = '')
	except:
		print("-- file name (", fileName, ") not found")
		return
	unload()
	csvreader = csv.reader(csvfile, delimiter=',')
	for row in csvreader:
		ss.contents.append(row)
		ss.numRows+=1

	ss.fileName = fileName
	ss.numCols = len(row)
	csvfile.close()
	printStats()


def ss_printrow(row):
	if row<0 or row>=ss.numRows:
		print("--row (",row,") is out of range")
		return
	print(' '.join(ss.contents[row]))


def ss_printrows(currentRow, endRow):
	if currentRow < 0:
		currentRow = 0
	while currentRow <= endRow and currentRow<ss.numRows:
		ss_printrow(currentRow)
		currentRow+=1


def ss_evalSum(col):
	sum = 0.0
	rowNum = 0
	for row in ss.contents:
		try:
			sum+= float(ss.contents[rowNum][col])
		except:
			sum+=0
		rowNum+=1
	return sum


def ss_evalAvg(col):
	return ss_evalSum(col)/ss.numRows


def ss_findRow(col, string, currentRow):
	if currentRow<0 or currentRow>=ss.numRows:
		print("-- start row", currentRow, ") is out of range")
		return
	while currentRow<ss.numRows:
		if string == ss.contents[currentRow][col]:
			print("-- found in row", currentRow)
			return
		currentRow+=1
	print("-- no matching row was found")


def ss_sortNumeric(col):
	if col<0 or col>=ss.numCols:
		print("-- column (", col, ") is out or range")
		return
	try:
		ss.contents.sort(key = lambda column: float(column[col]))
	except:
		pass


def ss_sort(col):
	if col<0 or col>=ss.numCols:
		print("-- column (", col, ") is out or range")
		return
	ss.contents.sort(key = lambda column: column[col])


def ss_save(fileName):
	csvfile = open(fileName, 'w', newline = '')
	csvwriter = csv.writer(csvfile, delimiter = ',')
	for row in ss.contents:
		csvwriter.writerow(row)
	csvfile.close()
	ss.fileName = fileName


def ss_merge(fileToMerge):
	csvfile = open(fileToMerge, 'r')
	csvreader = csv.reader(csvfile, delimiter=',')
	numCols = 0
	count = 0
	for row in csvreader:
		if count == 0:
			count+=1
			for col in row:
				numCols+=1
			if numCols != ss.numCols:
				print("-- can not merge: different number of columns in each file")
				return
		ss.contents.append(row)
		ss.numRows+=1


def ss_deleteRow(row):
	del(ss.contents[row])
	ss.numRows-=1


#############################################


def isFileLoaded():
	if ss.fileName == '':
		print("-- command ignored, no spreadsheet loaded")
		return False
	return True

def getColRowArg(command):
	try:
		num = int(command)
	except:
		if len(command) > 1:
			num = -1
		else:
			num = charToInt(command)
	return num


def charToInt(char):
	if char.isupper():
		num = ord(char) - ord('A')
	else:
		num = ord(char) - ord('a')
	if num >= 0 and num <= 25:
		return num
	return -1


def hasQuotes(command):
	if command[0] == '"' and command[-1] == '"':
		return command[1:-1]
	return command


def help():
	helptext = [
		"\nThe valid commands are:",
		"quit                     -- to exit the program",
		"help                     -- to display this help message",
		"load <filename>          -- to read in a spreadsheet",
		"save                     -- to save the spreadsheet back",
		"save <filename>          -- to save back to a different file",
		"merge <filename>         -- to read and append rows from another file",
		"stats                    -- to report on the spreadsheet size",
		"sort <col>               -- sort rows based on text data in column <col>",
		"sortnumeric <col>        -- sort rows based on numbers in column <col>",
		"deleterow <n>            -- delete row <n> from spreadsheet",
		"findrow <col> <text>     -- print the number of the first row",
		"                            such that column <col> holds <text>",
		"findrow <col> <text> <n> -- similarly, except search starts at",
		"                            row number n",
		"printrow <n>             -- print row number <n>",
		"printrow <n> <m>         -- print rows numbered <n> through <m>",
		"evalsum <col>            -- print the sum of the numbers that",
		"                            are in column <col>",
		"evalavg <col>            -- print the average of the numbers that",
		"                            are in column <col>"
		]
	for line in helptext:
		print(line)



#######################################################


def main():

	while True:
		print("\nEnter a subcommand ==> ", end = '')
		instring = input()
		command = instring.split()

		if instring == '' or command[0] == 'quit':
			unload()
			break;
		elif command[0] == 'help':
			help()
		elif command[0] == 'load':
			load(command)
		elif command[0] == 'printrow':
			printrow(command)
		elif command[0] == 'evalsum':
			evalSum(command)
		elif command[0] == 'evalavg':
			evalAvg(command)		
		elif command[0] == 'stats':
			printStats()
		elif command[0] == 'findrow':
			findRow(command)
		elif command[0] == 'sort':
			sort(command)
		elif command[0] == 'sortnumeric':
			sortNumeric(command)
		elif command[0] == 'save':
			save(command)
		elif command[0] == 'merge':
			merge(command)
		elif command[0] == 'deleterow':
			deleteRow(command)
		else:
			print("-- unrecongnized command (",instring, "). Type help for a list of commands")	
	print("Program exited")

################

main()

