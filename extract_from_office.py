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

import sys      # used for handling of command line options
import os       # used for reading of file system
import os.path
import getopt   # used for handling of command line options
from zipfile import ZipFile # used for handling of ZIP files
import re       # handling of regular expressions
import time     # to have sleep available

def check_extracted(folderpath='./input'):
    '''Checks if file was already extracted by looking for a corresponding folder name.'''

    folders = []
    files = []
    
    # Read the content of the watch folder
    with os.scandir(folderpath) as it:
        for entry in it:
            # Sort the entries into lists for folders and files
            if entry.is_dir():
                folders.append(entry)
            else:
                files.append(entry)
    
    # Check which files have not been extracted jet (those w/o related folders)
    for fl in files:
        # Remove the extension of the filename to get the related folder name
        folder, ext = os.path.splitext(fl)
        # If folder alread exists nothing has to be done
        if os.path.isdir(folder):
            print(fl, "already extracted")
        # Otherwise the file has to be extracted
        else:
            print(fl, "to be extracted")
    
    return(folders, files)
    
def main(argv):
    input_file = ''
    output_folder = ''
    
    try:
        opts, args = getopt.getopt(argv, "hi:w:o:",["ifile=", "ofolder=", "wfolder="])
    except getopt.GetoptError:
        print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
        print ('extract_from_office.py -w <watchfolder>')
        sys.exit(2)
    
    print(opts)    
    
    for opt, arg in opts:
        if opt == '-h':
            print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
            print ('extract_from_office.py -w <watchfolder>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input_file = arg
            print(opt, arg)
        elif opt in ("-o", "--ofolder"):
            output_folder = arg
            print(opt, arg)
        elif opt in ("-w", "--wfolder"):
            watch_folder = arg
            print(opt, arg)

    while watch_folder:
        print(check_extracted())
        time.sleep(1)
        
    # ~ # Create a ZipFile object and load file in it
    # ~ with ZipFile('input_file', 'r') as zipObj:
        # ~ # Get a list of all archived file names from the zip
        # ~ listOfFileNames = zipObj.namelist()
        # ~ # Iterate over the file names
        # ~ for fileName in listOfFileNames:
            # ~ print(fileName)
            # ~ # Check if file in dedicated folder or of requested filetype
            # ~ if re.search("(embeddings|media).+(bin|jpg|jpeg|doc|docx|ppt|pptx|pdf)", fileName):
                # ~ # Extract matching file
                # ~ zipObj.extract(fileName, 'output_folder')
                # ~ print("\t=>{}".format(fileName))
    return 0

# ~ if __name__ == '__main__':
    # ~ import sys
    # ~ sys.exit(main(sys.argv))
if __name__ == "__main__":
    main(sys.argv[1:])
