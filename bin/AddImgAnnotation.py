import time
import os
import sys
import re
import shutil
import ntpath
import exifread
import datetime as dt
import subprocess
from FolderTool import FolderTool


class AddImgAnnotation(FolderTool):
    def __init__(self, **kwargs):
        super(AddImgAnnotation, self).__init__()
        # Name the folder
        self.main_folder_name = kwargs.get('main_folder_name', 'img-add-annotation')
        self.out_folder_name  = kwargs.get('out_folder_name', '.')
        self.tool_name        = 'Add Annotation to Image'
        # Set patterns for matching files
        self.file_pattern=['^.*\.jpeg$','^.*\.jpe$', '^.*\.jpg$', '^.*\.png$', '^.*\.gif$']
        # get the folder paths and create not existing output-folders
        self.init_folders()
        # == Additional Properties ===========================================================
        self.path_convert = kwargs.get('path_convert', r'C:\PortableApps\cygwinx86\bin\convert.exe')
        self.param_set = list()
        self.set_annotation_parms()
        # Set Observer and Handler Actions
        self.in_handler.set_action_on_create(self.convert_and_overwrite_inputfile)
        self.in_observer.schedule(self.in_handler, path=self.main_folder_path, recursive=False)

        # Start Observing the Folders
        self.startObserve()

    def set_annotation_parms(self):
        self.param_set = list()
        self.param_set.append(
            {  'out': '.',
            'outExt': None,
             'inparms': [],
             'outparms': ['-fill \'#FFFFFFFF\'',  '-undercolor \'#00000080\'',  '-gravity Southwest', '-font Consolas', '-density 90', '-pointsize 40'],
             'annotate': '-annotate +0+5 \'{:%d.%m.%Y}\''})

    def convert_and_overwrite_inputfile(self, file_list=[]):
        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if
                         os.path.isfile(os.path.join(self.main_folder_path, name))]
        for fn in file_list:
            fn = ntpath.basename(fn)
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) and not self.is_temp_ignored(fn):
                fin = os.path.join(self.main_folder_path, fn)
                if self.wait_until_fileoperation_complete(fin):
                    for ps in self.param_set:
                        out_folder = os.path.join(self.main_folder_path, ps['out'])
                        if not os.path.isdir(out_folder):
                            os.makedirs(out_folder)
                        # Extract EXIF-TAGs
                        f = open(fin, 'rb')
                        tags = exifread.process_file(f)  # , stop_tag='EXIF DateTimeOriginal')
                        f.close()
                        if tags == {}:
                            tags['EXIF DateTimeOriginal'] = '1900:01:01 00:00:01'
                        # Convert EXIF Datetime to datetime object
                        dt_exif = dt.datetime.strptime("{:}".format(tags['EXIF DateTimeOriginal']), "%Y:%m:%d %H:%M:%S")

                        # Create Output File Name
                        fout = os.path.join(out_folder, fn)
                        fout = os.path.splitext(fout)[0]
                        if ps['outExt'] is None:
                            ps['outExt'] = os.path.splitext(fn)[1]

                        if 'outSuf' in ps.keys():
                            fout += ps['outSuf'] + '.'

                        fout += ps['outExt']
                        fout = os.path.abspath(fout)
                        # Create Annotate String
                        annot_str = ps['annotate'].format(dt_exif)
                        # Create Convert String
                        cmd_str = '"' + self.path_convert + '" ' + ' '.join(ps['inparms']) + ' "' + ntpath.basename(
                            fin) + '" ' + ' '.join(ps['outparms']) + ' ' + annot_str + ' "' + fout + '"'

                        # Ignore new File
                        self.add_temp_ignore(fout)

                        print('Annotating with command {:s}'.format(cmd_str))
                        # One conversion at a time
                        # subprocess.call(cmd_str)
                        # Multiple conversions at the same time
                        proc = subprocess.Popen(cmd_str, shell=False, stdin=None, stdout=None, stderr=None,
                                                close_fds=True,
                                                cwd=os.path.abspath(os.path.join(os.path.abspath(fin), os.path.pardir)))












if __name__ == '__main__':
    args = sys.argv[1:]
    ft = AddImgAnnotation()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()
