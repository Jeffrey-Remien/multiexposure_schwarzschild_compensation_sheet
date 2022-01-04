# Converts a Microsoft Excel 2007+ file into plain text
# for comparison using git diff
#
# Instructions for setup:
# 1. Place this file in a folder
# 2. Add the following line to the global .gitconfig:
#     [diff "zip"]
#   	    binary = True
#	    textconv = python c:/path/to/git_diff_xlsx.py
# 3. Add the following line to the repository's .gitattributes
#    *.xlsx diff=zip
# 4. Now, typing [git diff] at the prompt will produce text versions
# of Excel .xlsx files
#
# this readme and main function by William Usher
# Copyright William Usher 2013
# Contact: w.usher@ucl.ac.uk
#
# new parse function by Jeffrey Remien
# Copyright Jeffrey Remien 2021
# Contact: jeffrey.remien@gmail.comb
#

import pandas as pd
import sys

def parse(infile,outfile):
    """
    Converts an Excel file into text
    Returns a formatted text file for comparison using git diff.
    """
    pd.read_excel(infile).to_string(outfile, index=False)

# output cell address and contents of cell
def main():
    args = sys.argv[1:]
    if len(args) != 1:
        print ('usage: python git_diff_xlsx.py infile.xlsx')
        sys.exit(-1)
    outfile = sys.stdout
    parse(args[0],outfile)

if __name__ == '__main__':
    main()