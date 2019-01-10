# PYTHON script
# Purpose: Assigning material ID for Nastran profile
# Description: 1st we need to list the part name in A column of an excel sheet followed by MID1,MID2, MID3
#                   This script will take those data from excel and assign material ID as per respective part name

import os
import ansa
from ansa import utils, base, constants

def main():
	# File path of an excel and we've set the current directory to Nastran
	xl_element = utils.XlsxOpen('W:/Book1.xlsx')
	base.SetCurrentDeck(constants.NASTRAN)
	# To collect the information of the parts which are displayed
	partList = base.CollectEntities(constants.NASTRAN,None, "__PROPERTIES__", filter_visible=True)
	# To loop through each part, if matches the name in excel then assign respective MID1, MID2, MID3
	for nameList in range(0,len(partList)):
		for excel in range(0,len(partList)):
			firstColumn = "A"+str(excel+2)
			match = utils.XlsxGetCellValueByName(xl_element, "Sheet1", firstColumn)
			temp = partList[nameList]._name
			print(temp, type(temp))
			print(match, type(match))
			print(base.GetEntityCardValues(constants.NASTRAN, partList[nameList], ("Name", "MID1", "MID2", "MID3")))
			if temp ==  match:
				print(nameList +1)
				secondColumn = "B"+str(excel+2)
				thirdColumn = "C"+str(excel+2)
				fourthColumn = "D"+str(excel+2)
				value = utils.XlsxGetCellValueByName(xl_element, "Sheet1", secondColumn)
				value02 = utils.XlsxGetCellValueByName(xl_element, "Sheet1", thirdColumn)
				value03 = utils.XlsxGetCellValueByName(xl_element, "Sheet1", fourthColumn)
				base.SetEntityCardValues(constants.NASTRAN, partList[nameList], {"MID1":value, "MID2":value02, "MID3": value03})
	
if __name__ == '__main__':
	main()

