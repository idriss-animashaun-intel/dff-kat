import os
import urllib.request
import zipfile
import shutil
import time
from subprocess import Popen


dff_kat_master_directory = os.getcwd()
dff_kat_directory = dff_kat_master_directory+"\dff_kat-updates"
dff_kat_file = dff_kat_directory+"\\main\\main.exe"
dff_kat_rev = dff_kat_directory+"\\Rev.txt"


def installation():
    print("*** Downloading new version ***")
    urllib.request.urlretrieve("https://gitlab.devtools.intel.com/ianimash/dff_kat/-/archive/updates/dff_kat-updates.zip", dff_kat_master_directory+"\\dff_kat_new.zip")
    print("*** Extracting new version ***")
    zip_ref = zipfile.ZipFile(dff_kat_master_directory+"\dff_kat_new.zip", 'r')
    zip_ref.extractall(dff_kat_master_directory)
    zip_ref.close()
    os.remove(dff_kat_master_directory+"\dff_kat_new.zip")
    time.sleep(5)
    
def upgrade():    
    print("*** Removing old files ***")
    shutil.rmtree(dff_kat_directory)
    time.sleep(10)
    installation()


### Is dff_kat already installed? If yes get file size to compare for upgrade
if os.path.isfile(dff_kat_file):
    local_file_size = int(os.path.getsize(dff_kat_rev))
    # print(local_file_size)
    ### Check if update needed:
    f = urllib.request.urlopen("https://gitlab.devtools.intel.com/ianimash/dff_kat/-/raw/updates/Rev.txt") # points to the exe file for size
    i = f.info()
    web_file_size = int(i["Content-Length"])
    # print(web_file_size)


    if local_file_size != web_file_size:# upgrade available
        updt = input("*** New upgrade available! enter <y> to upgrade now, other key to skip upgrade *** ")
        if updt == "y": # proceed to upgrade
            upgrade()

### dff_kat wasn't installed, so we download and install it here                
else:
    install = input("Welcome to dff_kat! If you enter <y> dff_kat will be downloaded in the same folder where this file is.\nAfter the installation, this same file you are running now (\"dff_kat.exe\") will the one to use to open dff_kat :)\nEnter any other key to skip the download\n -->")
    if install == "y":
        installation()

print('Ready')


### We open the real application:
try:
    Popen(dff_kat_file)
    print("*** Opening DFF_KAT ***")
    time.sleep(20)
except:
    print('Failed to open application, Please open manually in subfolder')
    pass