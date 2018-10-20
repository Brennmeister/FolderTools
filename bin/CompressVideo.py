import time
import os
import sys
import re
import shutil
import ntpath
import subprocess
import threading
from FolderTool import FolderTool

class CompressVideo(FolderTool):
    def __init__(self, **kwargs):
        super(CompressVideo, self).__init__()
        # Name the folder
        self.main_folder_name = kwargs.get('main_folder_name', 'Compress Video')
        self.out_folder_name  = kwargs.get('out_folder_name', '.')
        self.tool_name        = 'Convert Video to Video with reduced Bitrate'
        # Set patterns for matching files
        self.file_pattern=kwargs.get('file_patterns', ['^.*\.mp4$'])
        # get the folder paths and create not existing output-folders
        self.init_folders()
        # == Additional Properties ===========================================================
        # Create a list to queue all the file-changes
        self.file_queue = list()
        self.ignore_time = 30*60 # 30 min ignore-time
        # create a timer object
        self.createTimer()

        self.path_ffmpeg = r"C:\PortableApps\ffmpeg-3.4.2-win64-static\bin\ffmpeg.exe"
        self.execAction = self.callConvert
        # Set Observer and Handler Actions
        self.in_handler.set_action_on_create(self.execAction)

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

    def callConvert(self, file_list=[]):
        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if
                         os.path.isfile(os.path.join(self.main_folder_path, name))]

        out_folder = os.path.join(self.out_folder_path)
        if not os.path.isdir(out_folder):
            os.makedirs(out_folder)

        for fn in file_list:
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, ntpath.basename(fn)) for x in self.ignore_files]):

                fout = os.path.join(out_folder, os.path.splitext(ntpath.basename(fn))[0]+'_crf25.mp4')
                # Create Command String
                cmd_str = '"{:s}" -loglevel panic -i "{:s}" -vcodec libx264 -crf 25 "{:s}"'.format(self.path_ffmpeg, fn, fout)

                # Add Out-File to ignore list. otherwise it would trigger new action
                self.ignore_files.append(ntpath.basename(fout))
                self.ignore_files_last_add = time.time()

                print('Converting with command {:s}'.format(cmd_str))
                # One conversion at a time
                subprocess.call(cmd_str)
                # Multiple conversions at the same time
                # proc = subprocess.Popen(cmd_str, shell=False, stdin=None, stdout=None, stderr=None,

if __name__ == '__main__':
    args = sys.argv[1:]
    ft = ConvertOffice2PDF()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()
