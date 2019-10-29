#!/usr/bin/python3
#
# Script   : extract_from_office.py
#
# Copyright: LiKneu 2019
#

# TODO: Switch -w allows to watch a directory and extract files as soon as a
#       file is placed in it (Ctrl-C stops the process)
# TODO: Add handling of command line arguments like -i input -o output etc.
# TODO: Add utilization of a config file
# TODO: Add output of useful information e.g. number of files, extracted ones

import sys
import getopt
from zipfile import ZipFile
import re

def main(argv):
    input_file = ''
    output_folder = ''
    
    try:
        opts, args = getopt.getopt(argv, "hi:o:",["ifile=", "ofolder="])
    except getopt.GetoptError:
        print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input_file = arg
        elif opt in ("-o", "--ofolder"):
            output_folder = arg
            
    # Create a ZipFile object and load file in it
    with ZipFile('input_file', 'r') as zipObj:
        # Get a list of all archived file names from the zip
        listOfFileNames = zipObj.namelist()
        # Iterate over the file names
        for fileName in listOfFileNames:
            print(fileName)
            # Check if file in dedicated folder or of requested filetype
            if re.search("(embeddings|media).+(bin|jpg|jpeg|doc|docx|ppt|pptx|pdf)", fileName):
                # Extract matching file
                zipObj.extract(fileName, 'output_folder')
                print("\t=>{}".format(fileName))
    return 0

# ~ if __name__ == '__main__':
    # ~ import sys
    # ~ sys.exit(main(sys.argv))
if __name__ == "__main__":
    main(sys.argv[1:])
