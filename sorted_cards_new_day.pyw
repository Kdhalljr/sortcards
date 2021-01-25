# copy to "xxxxxx"
import os
import time
from shutil import copyfile

date = time.strftime('%Y-%m-%d')
directory = 'Sorted Cards ' + date

if not os.path.exists(directory):
    os.makedirs(directory)

new_filepath = directory + '/' + 'Sorted Cards ' + date + '.xlsm'
if not os.path.exists(new_filepath):
    copyfile('sortcards script/SortCards.xlsm', new_filepath)

