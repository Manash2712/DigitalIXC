# This script expects 2 input from command line.
# 1 for file name/path from which data needs to be read.
# 2 for the string value based on which we need to select the data.

import pandas as pd
import re
import sys
import os

#This fun will use regex pattern to get the group names from rows
def match_pattern(strinp, variable):
    regex_pattern = fr'{variable} : \[code\]\<I\>(.*?)\</I\>\[/code\]'  #regex pattern for checking 
    match = re.search(regex_pattern, strinp)  #search for the occurence of pattern in strinp
    if match:
        group_data = match.group(1)  # Extract the data inside [code]<I> and </I>[/code]
        return group_data.split(',') #to get the multiple groups separated by comma
    else:
        return ''

#This fun will count the group occ and store it in dict
def count_group_occ(groups):
    grpMp = {}
    for grp in groups:
        grpMp[grp] = grpMp.get(grp, 0) + 1
    return grpMp

#This fun will create a output file and store the group names with their occurences in it 
def write_out_file(grpMp):
    with open('out_data.txt', 'w') as file:
        file.write(f"Group_name,Number of Occurences\n")
        # Loop through dictionary and write each key-value pair to the file
        for key, value in grpMp.items():
            file.write(f"{key},{value}\n")

#This fun will read the input file and get necessary data from it
def read_inp_file(fileName, stringToSearch):
    allGroups = []

    df = pd.read_excel(fileName, sheet_name="Input Data sheet")  #To read the file and sheet
    name_column = df['Additional comments'].tolist()  #To read the specific column with name Additional comments and store the column data in list
    
    for item in name_column:  #iterating through the column data row wise
        valueNames = match_pattern(item, stringToSearch)  #to get the groupnames
        for value in valueNames:
            allGroups.append(value)  #To store grp names individually into all groups list

    grpMp = count_group_occ(allGroups)  #count each group occurences
    write_out_file(grpMp)  #write the output to file

def main():
    if len(sys.argv) == 3:
        print(sys.argv[0])
        inpFile = sys.argv[1]
        stringToSearch = sys.argv[2]
        if os.path.isfile(inpFile):
            read_inp_file(inpFile, stringToSearch)
        else:
            print('File does not exist')
    else:
        print("Please provide file path/name and string to be searched in file")

if __name__ == "__main__":
    print(__name__)
    main()