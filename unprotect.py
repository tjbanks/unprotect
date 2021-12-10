import glob
import os
import pathlib
import re
import shutil
import sys
import time
import zipfile

def remove_protection(file_path):

    file_extension = pathlib.Path(file_path).suffix

    # rename the file to .zip
    file_basename = os.path.basename(file_path)
    zip_path = re.sub(file_extension +'$','.zip',file_path)
    temp_dir = os.path.dirname(zip_path)
    if not os.path.exists(zip_path):
        os.rename(file_path,zip_path)

    # extract zip
    extracted_path = os.path.join(temp_dir,'extracted')
    if not os.path.exists(extracted_path):
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_path)

    if file_extension == '.xlsx':
        # get all sheets xml files to edit
        sheets_dir = os.path.join(extracted_path, 'xl', 'worksheets')
        sheets_glob = os.path.join(sheets_dir,'*.xml')
        sheets = glob.glob(sheets_glob)

        for sheet in sheets:
            with open(sheet, 'r+') as f:
                text = f.read()
                # remove protection
                text = re.sub('<sheetProtection.*?\/>','',text)
                f.seek(0)
                f.write(text)
                f.truncate()
    
    elif file_extension == '.docx':
        settings_xml = os.path.join(extracted_path,'word','settings.xml')
        with open(settings_xml, 'r+') as f:
            text = f.read()
            # remove protection
            text = re.sub('<w:documentProtection.*?\/>','',text)
            f.seek(0)
            f.write(text)
            f.truncate()
     

    # zip it all back up
    new_zip_path = re.sub('.zip$','_unprotected.zip',zip_path)
    prefix = re.sub('.zip$','',new_zip_path)
    cwd = os.getcwd()
    
    with zipfile.ZipFile(new_zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        info = zipfile.ZipInfo(prefix+'\\')
        zipf.writestr(info, '')
        os.chdir(extracted_path)
        for root, dirs, files in os.walk('./'):
            for d in dirs:
                info = zipfile.ZipInfo(os.path.join(root, d)+'\\')
                zipf.writestr(info, '')
            for file in files:
                zipf.write(os.path.join(root, file))
    os.chdir(cwd)

    # turn back into original file extension
    new_file_path = re.sub('.zip$',file_extension,new_zip_path)
    if os.path.exists(new_zip_path):
        shutil.copy(new_zip_path,new_file_path)

    return new_file_path

def run(file_path=None):
    
    temp_path = './unprotect.temp'

    if file_path is None:
        print("No input file specified. Drag the protected Excel or Word document onto this executable.\n")
        print("The resulting unprotected xlsx or docx file will be located in the original file directory.\n\n")
        input('Press enter to exit...')
        return

    # create a temp path for working in
    if not os.path.exists(temp_path):
        os.makedirs(temp_path)

    # copy the file to destination dir
    dst_file = os.path.join(temp_path,os.path.basename(file_path))
    if not os.path.exists(dst_file):
        shutil.copy(file_path,temp_path) 
    
    # remove password
    unprotected_file = remove_protection(dst_file)
    shutil.copy(unprotected_file,'./')

    # remove the temporary path
    if os.path.exists(temp_path) and os.path.isdir(temp_path):
        shutil.rmtree(temp_path)
        

if __name__ == '__main__':
    if __file__ != sys.argv[-1]:
        run(sys.argv[-1])
    else:
        run(None)