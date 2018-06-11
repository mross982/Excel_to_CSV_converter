#!/usr/bin/env python

#************************************************************************************
#  - Process:
#       - Filters Excel files from given directory by given file extension and (optionally) filename matching criteria
#       - Converts selected Excel files to .csv format
#       - Relocates converted csv files to given archive directory
#       - Optionally, deletes the original Excel files
#
#   - Author: Russell Ragosta
#   - Date 05/29/2018
#   - Last Modified:
#   - Modification:
#
#**********************************************************************************

import sys
import glob
import pandas as pd
import shutil
import os
import logging
import time
import xlrd

class TASKS(object):
    '''
    Each task in the list below contains 5 pieces of information that MUST be in the correct order: 1. Directory from which the Excel files are currently located; 2. Directory which you want the csv converted files to be located; 3. The file extension of the Excel files you wish to convert and relocate; 4. Either a string to match the Excel filenames OR False if you want convert all Excel files; and 5. True to delete the original Excel files OR False to keep the original Excel files.

    '''
    tasks = [

    ]


    #******** TESTING PURPOSES ONLY *****************************
    test_dir = os.path.dirname(os.path.realpath(__file__))

    test_tasks = [
    [os.path.join(test_dir, 'directory_1'), os.path.join(test_dir, 'directory_2'), '.xlsx', 'example_', True],
    [os.path.join(test_dir, 'directory_1'), os.path.join(test_dir, 'directory_2'), '.xlsx', False, True]
    ]


class obj(object):

    def __init__(self, dir1, dir2, xlfiletype, match=False, delete_originals=False):
        self.primary_filepath = dir1
        self.secondary_filepath = dir2
        self.file_type = xlfiletype
        self.match = match
        self.delete_originals = delete_originals


    def seek_type(self):
        '''
        Filters through files in a given directory and returns a list of filenames of a given file type
        '''

        ls_fromFiles = []
        ls_allFiles = os.listdir(self.primary_filepath)
        for f in ls_allFiles:
            if f.endswith(self.file_type):
                ls_fromFiles.append(f)

        if self.match != False:
            ls_fromFiles = obj.seek_match(self, ls_fromFiles)

        obj.convert_n_archive(self, ls_fromFiles)


    def seek_match(self, ls_files):
        '''
        Filters through files in a given list and returns a list of filenames that match a given substring
        @ ls_files (list) - a list of file names to filter against the substring match criteria
        '''

        ls_fromFiles = []
        for f in ls_files:
            if self.match in f:
                ls_fromFiles.append(f)

        return ls_fromFiles


    def convert_n_archive(self, ls_files):
        '''
        Converts a list of files to csv format, saves to new directory, then deletes original file if that option is selected
        @ ls_files (list) - a list of file names to be converted
        '''

        for file in ls_files:
            in_filepath = os.path.join(self.primary_filepath, file)
            df = pd.read_excel(in_filepath)
            base = os.path.splitext(file)[0]
            out_filepath = os.path.join(self.secondary_filepath, base + '.csv')
            df.to_csv(out_filepath, index=False)
            if self.delete_originals == True:
                os.remove(in_filepath)


if __name__ == '__main__':
    for file in TASKS.test_tasks:
        dir1 = file[0]
        dir2 = file[1]
        xlfiletype = file[2]
        match = file[3]
        delete_originals = file[4]
        myobj = obj(dir1, dir2, xlfiletype, match, delete_originals)

        myobj.seek_type()

