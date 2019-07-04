import time
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler

import os
import sys
import pathlib

from AddEXIFDatePrefix import  AddEXIFDatePrefix
from FileRenameAndNumber import FileRenameAndNumber
from ConvertImage import ConvertImage
from ConvertPDF import ConvertPDF
from ConvertOffice2PDF import ConvertOffice2PDF
from AddImgAnnotation import AddImgAnnotation

# Desktop
cfg = {'path_imconvert': r'D:\cygwin\bin\convert.exe',
       'path_qpdf': r'C:\Users\Manuel\Documents\FolderTools\lib\qpdf-8.1.0\bin\qpdf.exe'
      }
# Notebook
if os.path.exists(r'C:\PortableApps') and not os.path.exists(cfg['path_qpdf']):
    cfg = {'path_imconvert': r'C:\PortableApps\cygwinx86\bin\convert.exe',
           'path_qpdf': r'C:\PortableApps\cygwinx86\bin\qpdf.exe'
          }


if __name__ == '__main__':
    tools = list()
    tools.append(FileRenameAndNumber())
    tools.append(AddEXIFDatePrefix(main_folder_name='img-add-dateprefix-Y-m',   prefix_format='{:%Y-%m} - '))
    tools.append(AddEXIFDatePrefix(main_folder_name='img-add-dateprefix-Y-m-d', prefix_format='{:%Y-%m-%d} - '))
    tools.append(AddEXIFDatePrefix(main_folder_name='img-exif-doublette', prefix_format='{:%Y-%m-%d_%H%M%S} {:s} {:s}', append_original_file_name=False, prevent_overwrite=False))

    tools.append(ConvertImage(main_folder_name='img-to-pdf', preset_parms='pdf', path_convert=cfg['path_imconvert']))
    tools.append(ConvertImage(main_folder_name='pdf-to-png', preset_parms='pdf2png', path_convert=cfg['path_imconvert']))
    tools.append(ConvertImage(main_folder_name='img-to-thumb', preset_parms='thumb', path_convert=cfg['path_imconvert']))
    tools.append(ConvertImage(main_folder_name='img-to-Q70', preset_parms='quality70', path_convert=cfg['path_imconvert']))
    tools.append(AddImgAnnotation(main_folder_name='img-add-annotation', path_convert=cfg['path_imconvert']))

    tools.append(ConvertPDF(main_folder_name='pdf-merge', pdf_action='merge', path_qpdf=cfg['path_qpdf']))
    tools.append(ConvertPDF(main_folder_name='pdf-split', pdf_action='split', path_qpdf=cfg['path_qpdf']))
    tools.append(ConvertPDF(main_folder_name='pdf-decrypt', pdf_action='decrypt', path_qpdf=cfg['path_qpdf']))

    tools.append(ConvertOffice2PDF(main_folder_name='PPT2PDF', file_patterns=['^.*\.ppt$', '^.*\.pptx$'], preset='ppt2pdf'))
    tools.append(ConvertOffice2PDF(main_folder_name='DOC2PDF', file_patterns=['^.*\.doc$', '^.*\.docx$'], preset='doc2pdf'))

    for ft in tools:
        print('Tool started: {:s}'.format(ft.tool_name))

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        for ft in tools:
            ft.stopObserve()