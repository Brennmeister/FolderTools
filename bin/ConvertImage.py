import time
import os
import sys
import re
import shutil
import ntpath
import subprocess
from FolderTool import FolderTool

class ConvertImage(FolderTool):
    def __init__(self, **kwargs):
        super(ConvertImage, self).__init__()
        # Name the folder
        self.main_folder_name = kwargs.get('main_folder_name', 'img-to-pdf')
        self.out_folder_name  = kwargs.get('out_folder_name', '.')
        self.tool_name        = 'Image conversion'
        # Set patterns for matching files
        self.file_pattern=['^.*\.jpeg$','^.*\.jpe$', '^.*\.jpg$', '^.*\.png$']
        # get the folder paths and create not existing output-folders
        self.init_folders()
        # == Additional Properties ===========================================================
        self.path_convert = kwargs.get('path_convert', r'C:\PortableApps\cygwinx86\bin\convert.exe')
        self.param_set = list()
        if kwargs.get('preset_parms') == 'pdf':
            self.set_pdf_parms()
        elif kwargs.get('preset_parms') == 'thumb':
            self.set_thumb_parms()
        elif kwargs.get('preset_parms') == 'pdf2png':
             self.set_pdf2png_parms()
        # Set Observer and Handler Actions
        self.in_handler.set_action_on_create(self.convert_and_keep_inputfile)
        self.in_observer.schedule(self.in_handler, path=self.main_folder_path, recursive=False)

        # Start Observing the Folders
        self.startObserve()

    def set_pdf_parms(self):
        self.param_set = list()
        self.param_set.append(
            {  'out': 'HQ',
            'outExt': '.pdf',
             'inparms': [],
             'outparms': ['-density 150', '-quality 100', '-trim', '-sharpen 0x2.0', '-level 0%,100%']})
        self.param_set.append(
            {  'out': 'MQ',
            'outExt': '.pdf',
             'inparms': [],
             'outparms': ['-density 75', '-quality 50', '-trim', '-sharpen 0x2.0', '-level 0%,100%']})
        self.param_set.append(
            {  'out': 'MQ_text3070',
            'outExt': '.pdf',
             'inparms': [],
             'outparms': ['-density 75', '-quality 50', '-trim', '-sharpen 0x2.0', '-level 30%,70%']})
    def set_thumb_parms(self):
        self.param_set = list()
        self.param_set.append(
            {  'out': 'Q70 2000x',
            'outExt': '.jpg',
             'inparms': [],
             'outparms': ['-quality 70', '-geometry 2000x']})
        self.param_set.append(
            {  'out': 'Q70 1000x',
            'outExt': '.jpg',
             'inparms': [],
             'outparms': ['-quality 70', '-geometry 1000x']})
        self.param_set.append(
            {  'out': 'Q70 500x',
            'outExt': '.jpg',
             'inparms': [],
             'outparms': ['-quality 70', '-geometry 500x']})
    def set_pdf2png_parms(self):
        self.param_set = list()
        self.param_set.append(
            {     'out': '.',
               'outExt': '.png',
               'outSuf': '_page%03d',
              'inparms': ['-density 150'],
             'outparms': ['-resize 100%', '-quality 100']})
        self.file_pattern = ['^.*\.pdf$']
    def convert_and_keep_inputfile(self, file_list=[]):
        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if
                         os.path.isfile(os.path.join(self.main_folder_path, name))]
        for fn in file_list:
            fn = ntpath.basename(fn)
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, fn) for x in self.ignore_files]):
                fin = os.path.join(self.main_folder_path, fn)
                if self.wait_until_fileoperation_complete(fin):
                    for ps in self.param_set:
                        out_folder = os.path.join(self.main_folder_path, ps['out'])
                        if not os.path.isdir(out_folder):
                            os.makedirs(out_folder)
                        # Create Output File Name
                        fout = os.path.join(out_folder, fn)
                        fout = os.path.splitext(fout)[0]
                        if ps['outExt'] is None:
                            ps['outExt'] = os.path.splitext(fn)[1]

                        if 'outSuf' in ps.keys():
                            fout += ps['outSuf']

                        fout += ps['outExt']
                        fout = os.path.abspath(fout)
                        # Create Convert String
                        cmd_str = '"' + self.path_convert + '" ' + ' '.join(ps['inparms']) + ' "' + ntpath.basename(fin) + '" ' + ' '.join(ps['outparms']) + ' "' + fout + '"'
                        print('Converting with command {:s}'.format(cmd_str))
                        # One conversion at a time
                        # subprocess.call(cmd_str)
                        # Multiple conversions at the same time
                        proc = subprocess.Popen(cmd_str, shell=False, stdin=None, stdout=None, stderr=None,
                                                close_fds=True, cwd=os.path.abspath(os.path.join(os.path.abspath(fin), os.path.pardir)))


if __name__ == '__main__':
    args = sys.argv[1:]
    ft = ConvertImage()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()
