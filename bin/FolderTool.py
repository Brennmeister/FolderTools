import time
import os
import sys
import re
import shutil
import ntpath

from watchdog.observers import Observer
from MyHandler import MyHandler

class FolderTool:
    def __init__(self):
        # Name the folder
        self.main_folder_name = '.'
        self.out_folder_name  = '.'
        self.tool_name        = ''

        # Set valid file patterns
        self.file_pattern=['.*']

        # Store for saving file-names to ignore
        self.ignore_files = list()
        self.ignore_files_last_add = 0

        # Create Objects for in and out handler
        self.in_observer  = Observer()
        self.in_handler   = MyHandler()
        self.out_observer = Observer()
        self.out_handler  = MyHandler()

        # Define duration for temporary ignore
        self.ignore_time = 1 # Ignore for 1 s

    def init_folders(self):
        # get the folder paths
        self.script_file_path = os.path.realpath(__file__)
        self.main_folder_path = os.path.realpath(
            os.path.join(os.path.dirname(self.script_file_path), '../', self.main_folder_name))
        self.out_folder_path  = os.path.realpath(
            os.path.join(os.path.dirname(self.script_file_path), '../', self.main_folder_name, self.out_folder_name))

        # create Input-folder if it does not exist
        if not os.path.isdir(self.main_folder_path):
            os.makedirs(self.main_folder_path)
        # create Output-folder if it does not exist
        if not os.path.isdir(self.out_folder_path):
            os.makedirs(self.out_folder_path)

    def get_outfile_count(self):
        return len([name for name in os.listdir(self.out_folder_path) if os.path.isfile(name)])


    def wait_until_fileoperation_complete(self, fin,  max_tries=10):
        while max_tries > 0:
            max_tries -= 1
            try:
                file = open(fin, 'r+')
                file.close()
                return True
            except (PermissionError, IOError) as e:
                if max_tries > 0:
                    time.sleep(0.1)  # unfortunately the pause is needed :(
                else:
                    raise e

    def ft_safe_move(self, fin, fout, prevent_overwrite=True):
        # Moves the file fin to fout after waiting for the file to be complete
        # Also add the moved file to the ignore-list
        if self.wait_until_fileoperation_complete(fin):
            # Make sure the output-File does not exist yet
            i=0
            fout_new=fout
            if prevent_overwrite:
                while os.path.isfile(fout_new):
                    (path, fn) = ntpath.split(fout)
                    (fn, ext) = ntpath.splitext(fn)
                    fn = '{:s}_{:02.0f}'.format(fn,i)
                    fout_new=os.path.join(path, fn+ext)
                    i+=1
                fout = fout_new

            # Move input file to output file
            self.add_temp_ignore(fout)
            shutil.move(fin, fout)
            print('shutil.move({:s}, {:s})'.format(fin, fout))

            return fout

    def add_temp_ignore(self, fn):
        self.ignore_files.append(ntpath.basename(fn))
        self.ignore_files_last_add = time.time()

    def is_temp_ignored(self,fn):
        # Check if time for temp ignore is already up
        if  time.time() > self.ignore_files_last_add + self.ignore_time:
            # time is up, so delete ignored files
            self.ignore_files = []

        if ntpath.basename(fn) in self.ignore_files:
            return True
        else:
            return False

    def startObserve(self):
        self.in_observer.start()
        self.out_observer.start()

    def stopObserve(self):
        self.in_observer.stop()
        self.out_observer.stop()

        self.in_observer.join()
        self.out_observer.join()

