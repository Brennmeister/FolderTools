import time
import os
import sys
import re
import shutil
import ntpath

from FolderTool import FolderTool

class FileRenameAndNumber(FolderTool):
    def __init__(self):
        super(FileRenameAndNumber, self).__init__()

        # Name the folder
        self.main_folder_name = 'file-rename-and-number'
        self.out_folder_name  = '.'
        self.tool_name        = 'File rename and renumber'
        # get the folder paths and create not existing output-folders
        self.init_folders()

        # == Additional Properties ===========================================================
        # Set the Prefix for the renaming
        self.prefix_format = '{:03.0f} - '
        # Set the next number for the renaming
        self.autoset_next_number()
        self.file_pattern=['.*']

        # Set Observer and Handler Actions
        self.in_handler.set_action_on_create(self.rename_and_move)
        self.in_handler.set_action_on_delete(self.autoset_next_number)
        self.in_observer.schedule(self.in_handler, path=self.main_folder_path, recursive=False)

        # Start Observing the Folders
        self.startObserve()

    def autoset_next_number(self):
        self.next_number = self.get_outfile_count() + 1


    def rename_and_move(self, file_list=[]):
        # update the ignored files
        if (time.time() - self.ignore_files_last_add)>1:
            # Clear list of timestamp for last added files
            self.ignore_files=list()

        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if os.path.isfile(os.path.join(self.main_folder_path, name))]
        for fn in file_list:
            fn = ntpath.basename(fn)
            # File must be in valide extensions and file must not be in ignore-list
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, fn) for x in self.ignore_files]):

                fin = os.path.join(self.main_folder_path, fn)
                while True: # Generate new numbers until one is unique. Should not be needed in normal operation
                    fout = os.path.join(self.out_folder_path, '{:s}{:s}'.format(self.prefix_format.format(self.next_number), fn))
                    self.next_number += 1
                    if not os.path.isfile(fout):
                        break

                # Try to handle Errors due to not fully written files
                self.ft_safe_move(fin, fout)


if __name__ == '__main__':
    args = sys.argv[1:]
    ft = FileRenameAndNumber()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()