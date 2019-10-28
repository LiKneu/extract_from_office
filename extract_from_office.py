#!/usr/bin/python3
#
# Script   : extract_from_office.py
#
# Copyright: LiKneu 2019
#

from zipfile import ZipFile
import re

def main(args):
    # Create a ZipFile object and load file in it
    with ZipFile('./input/FAI-1HDG911412P0002( Gear Wheel).xlsx', 'r') as zipObj:
        # Get a list of all archived file names from the zip
        listOfFileNames = zipObj.namelist()
        # Iterate over the file names
        for fileName in listOfFileNames:
            print(fileName)
            # Check if file in dedicated folder or of requested filetype
            if re.search("(embeddings|media).+(bin|jpg|jpeg|doc|docx|ppt|pptx|pdf)", fileName):
                # Extract matching file
                zipObj.extract(fileName, 'output')
                print("\t=>{}".format(fileName))
    return 0

if __name__ == '__main__':
    import sys
    sys.exit(main(sys.argv))
