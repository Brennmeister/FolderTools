import time
import os
import sys
import re
import shutil
import ntpath
import exifread
import datetime as dt
from FolderTool import FolderTool

class AddEXIFDatePrefix(FolderTool):
    def __init__(self, **kwargs):
        super(AddEXIFDatePrefix, self).__init__()
        # Name the folder
        self.main_folder_name = kwargs.get('main_folder_name', 'img-add-dateprefix-Y-m')
        self.out_folder_name  = kwargs.get('out_folder_name','.')
        self.tool_name        = 'Add exif-date to image'
        # Set patterns for matching files
        self.file_pattern     = kwargs.get('file_pattern', ['^.*\.jpeg$','^.*\.jpe$', '^.*\.jpg$'])

        # get the folder paths and create not existing output-folders
        self.init_folders()

        # == Additional Properties ===========================================================
        # Set the Prefix for the renaming
        self.prefix_format = kwargs.get('prefix_format','{:%Y-%m} - ')
        self.append_original_file_name = kwargs.get('append_original_file_name', True)
        self.prevent_overwrite = kwargs.get('prevent_overwrite', True)

        # Set Observer and Handler Actions
        self.in_handler.set_action_on_create(self.renameImage)
        self.in_observer.schedule(self.in_handler, path=self.main_folder_path, recursive=False)

        # Start Observing the Folders
        self.startObserve()

    def renameImage(self, file_list=[]):
        # update the ignored files
        if (time.time() - self.ignore_files_last_add)>1:
            # Clear list of timestamp for last added files
            self.ignore_files=list()

        if file_list is None:
            file_list = [name for name in os.listdir(self.main_folder_path) if os.path.isfile(os.path.join(self.main_folder_path, name))]
        for fn in file_list:
            fn = ntpath.basename(fn)
            # File must be in valid extensions and file must not be in ignore-list
            if any([re.match(x, fn, re.IGNORECASE) for x in self.file_pattern]) \
                    and not any([re.match(x, fn) for x in self.ignore_files]):
                fin = os.path.join(self.main_folder_path, fn)
                out_folder_path = self.out_folder_path
                if not os.path.isdir(out_folder_path):
                    os.makedirs(out_folder_path)

                if self.wait_until_fileoperation_complete(fin):
                    # Extract Tags
                    f = open(fin,'rb')
                    tags = exifread.process_file(f)#, stop_tag='EXIF DateTimeOriginal')
                    f.close()

                    temp_prevent_overwrite = False
                    has_exif_data = True
                    if tags=={} or 'EXIF DateTimeOriginal' not in tags.keys():
                        tags['EXIF DateTimeOriginal']='1900:01:01 00:00:01'
                        temp_prevent_overwrite = True
                        has_exif_data = False
                    # Add missing infos
                    if 'Image Make' not in tags.keys():
                        tags['Image Make'] = 'na'
                    else:
                        tags['Image Make'] = '{:}'.format(tags['Image Make'])
                    if 'Image Model' not in tags.keys():
                        tags['Image Model'] = 'na'
                    else:
                        tags['Image Model'] = '{:}'.format(tags['Image Model'])


                dt_exif = dt.datetime.strptime("{:}".format(tags['EXIF DateTimeOriginal']), "%Y:%m:%d %H:%M:%S")
                fnn = os.path.splitext(fn)[0]
                fout = os.path.join(out_folder_path, '{:s}{:s}.jpg'.format(self.prefix_format, fnn if self.append_original_file_name else '').format(dt_exif, tags['Image Make'], tags['Image Model']))

                if has_exif_data:
                    self.ft_safe_move(fin,fout, self.prevent_overwrite+temp_prevent_overwrite)





if __name__ == '__main__':
    args = sys.argv[1:]
    ft = AddEXIFDatePrefix()

    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        ft.stopObserve()

    ft.stopObserve()
