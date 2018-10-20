import time
import os
import sys
import re
import shutil
import ntpath
import subprocess
import threading
from FolderTool import FolderTool

class ConvertPDF(FolderTool):
    def __init__(self, **kwargs):
        super(ConvertPDF, self).__init__()
        # Name the folder
        self.main_folder_name = kwargs.get('main_folder_name', 'pdf-merge')
        self.out_folder_name  = kwargs.get('out_folder_name', '.')
        self.tool_name        = 'Merge multiple PDFs/Split PDF'
        # Set patterns for matching files
        self.file_pattern=['^.*\.pdf$']
        # get the folder paths and create not existing output-folders
        self.init_folders()
        # == Additional Properties ===========================================================
        self.path_qpdf = kwargs.get('path_qpdf', r'C:\PortableApps\cygwinx86\bin\qpdf.exe')
        # Create a list to queue all the file-changes
        self.file_queue = list()
        # create a timer object
        self.createTimer()
        pdf_action = kwargs.get('pdf_action', 'merge')
        if pdf_action=='merge':
            self.execAction = self.mergePDF
            # Set Observer and Handler Actions
            self.in_handler.set_action_on_create(self.collectChanges)
        elif pdf_action=='decrypt':
            self.execAction = self.decryptPDF
            # Set Observer and Handler Actions
            self.in_handler.set_action_on_create(self.decryptPDF)


        #self.in_handler.set_action_on_create(self.splitPDF)
        self.in_observer.schedule(self.in_handler, path=self.main_folder_path, recursive=False)

        # Start Observing the Folders
        self.startObserve()

    def createTimer(self):
        self.timer = threading.Timer(0.1, self.execQueue)

    def restartTimer(self):
        self.timer.cancel()
        self.createTimer()
        self.timer.start()

    def collectChanges(self,fin_list):
        for fin in fin_list:
            if ntpath.basename(fin) not in self.ignore_files:
                self.timer.cancel()
                # Collects the files changed/created to queue
                self.wait_until_fileoperation_complete(fin)
                self.file_queue.append(fin)
                # reset timer until mergePDF is fired
                self.restartTimer()

    def execQueue(self):
        print('files in queue:')
        for f in self.file_queue:
            print('{:s}, '.format(f))
        self.execAction(self.file_queue)
        self.file_queue=list()

    def mergePDF(self, file_list=[]):
        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if
                         os.path.isfile(os.path.join(self.main_folder_path, name))]
        cmd_in_parms = list()
        for fn in file_list:
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, ntpath.basename(fn)) for x in self.ignore_files]):
                cmd_in_parms.append('"{:s}" 1-z'.format(fn))

        out_folder = os.path.join(self.out_folder_path)
        if not os.path.isdir(out_folder):
            os.makedirs(out_folder)

        fout = os.path.join(out_folder, 'merged_'+ntpath.basename(fn))
        # Create Command String
        cmd_str = '"' + self.path_qpdf + '" --empty --pages ' + ' '.join(cmd_in_parms) + ' -- "' + fout + '"'

        # Add Out-File to ignore list. otherwise it would trigger new action
        self.ignore_files.append(ntpath.basename(fout))
        self.ignore_files_last_add = time.time()

        print('Converting with command {:s}'.format(cmd_str))
        # One conversion at a time
        # subprocess.call(cmd_str)
        # Multiple conversions at the same time
        proc = subprocess.Popen(cmd_str, shell=False, stdin=None, stdout=None, stderr=None,
                                close_fds=True)

    def decryptPDF(self, file_list=[]):
        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if
                         os.path.isfile(os.path.join(self.main_folder_path, name))]
        for fn in file_list:
            out_folder = os.path.join(self.out_folder_path)
            if not os.path.isdir(out_folder):
                os.makedirs(out_folder)
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, ntpath.basename(fn)) for x in self.ignore_files]):

                fout = os.path.join(out_folder, 'dec_'+ntpath.basename(fn))
                # Create Command String
                cmd_str = '"' + self.path_qpdf + '" --decrypt ' + ' "'+fn+ '" "' + fout + '"'

                # Add Out-File to ignore list. otherwise it would trigger new action
                self.ignore_files.append(ntpath.basename(fout))
                self.ignore_files_last_add = time.time()

                print('Converting with command {:s}'.format(cmd_str))
                # One conversion at a time
                # subprocess.call(cmd_str)
                # Multiple conversions at the same time
                proc = subprocess.Popen(cmd_str, shell=False, stdin=None, stdout=None, stderr=None,
                                        close_fds=True)

if __name__ == '__main__':
    args = sys.argv[1:]
    ft = ConvertPDF()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()
