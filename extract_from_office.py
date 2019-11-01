#!/usr/bin/python3
#
# Script   : extract_from_office.py
#
# Copyright: LiKneu 2019
#

# TODO: Add utilization of a config file

import sys      # used for handling of command line options
import os       # used for reading of file system
import os.path
import getopt   # used for handling of command line options
import zipfile
from zipfile import ZipFile # used for handling of ZIP files
import re       # handling of regular expressions
import time     # to have sleep available

# Define ANSI escape sequences to allow colorful output in the terminal.
# Code was taken from:
# https://stackoverflow.com/questions/287871/how-to-print-colored-text-in-terminal-in-python
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    
def extract_files(input_file, output_folder):
    '''Extract files embedded in given MS Office document.'''
    # This code was taken from:
    # https://bitdrop.st0w.com/2010/07/23/python-extracting-a-file-from-a-zip-file-with-a-different-name/
    # It allows to extract a file and rename it at the same time
    if not zipfile.is_zipfile(input_file):
        print(bcolors.WARNING + "\n{} not a zipfile.".format(input_file))
        print("Please remove this file from watchfolder.")
        input("Press ENTER to continue:" + bcolors.ENDC)
        return 1
    
    zipdata = ZipFile(input_file)
    zipinfos = zipdata.infolist()
    
    filecount = 0   # Counte for the number of extracted files
    
    print("\nStart extraction of embedded files...")
    
    for zipinfo in zipinfos:
        # Check if file in dedicated folder or of requested filetype
        if re.search("(embeddings|media).+(bin|jpg|jpeg|doc|docx|xls|xlsx|ppt|pptx|pdf)", zipinfo.filename):
            # Rename files with extension .bin to .pdf
            zipinfo.filename = re.sub('\.bin$', '.pdf', zipinfo.filename)
            # Extract the file
            zipdata.extract(zipinfo, output_folder)
            filecount += 1
            print("\t=>{}".format(zipinfo.filename))
    
    if not filecount:
        print("\tNo embedded files found to extract.")
    
    print("Extracted {} of {} files.".format(filecount, len(zipinfos)))
    
    return 0
                
def check_extracted(folderpath):
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
        # If extracted folder doesn't exist jet it has to be extracted now
        if not os.path.isdir(folder):
            print(fl, "to be extracted")
            extract_files(fl, folder)
            print(bcolors.OKGREEN + "\nRunning in watchdog mode. <Ctrl+c> ends program.\n" + bcolors.ENDC)
    return 0
    
def main(argv):
    input_file = ''
    output_folder = ''
    watch_folder = ''
    
    try:
        opts, args = getopt.getopt(argv, "hi:w:o:",["ifile=", "ofolder=", "wfolder="])
    except getopt.GetoptError:
        print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
        print ('extract_from_office.py -w <watchfolder>')
        sys.exit(2)
    
    for opt, arg in opts:
        if opt == '-h':
            print ('extract_from_office.py -i <inputfile> -o <outputfolder>')
            print ('extract_from_office.py -w <watchfolder>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input_file = arg
        elif opt in ("-o", "--ofolder"):
            output_folder = arg
        elif opt in ("-w", "--wfolder"):
            watch_folder = arg

    if watch_folder:
        print(bcolors.OKGREEN + "\nRunning in watchdog mode. <Ctrl+c> ends program.\n" + bcolors.ENDC)
        while watch_folder:
            check_extracted(watch_folder)
            time.sleep(2)
    elif input_file and output_folder:
        extract_files(input_file, output_folder)
    
    return 0

if __name__ == "__main__":
    main(sys.argv[1:])
