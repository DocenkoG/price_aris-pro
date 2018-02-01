# -*- coding: UTF-8 -*-
import logging
import logging.config
import configparser
import shutil
from aris_pro_converter_cables import convert2csv_cables 
from aris_pro_converter_pro    import convert2csv_pro 
from aris_pro_converter_dsp    import convert2csv_dsp 
from aris_pro_converter_pa     import convert2csv_pa 
import re
import os


def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def convert2csv( pFileName   # file for convertation
               , myname      # organisatiom name (dir name)
               ) :
    global log
    global SheetName
    global FilenameIn
    global FilenameOut
    make_loger()
    log.debug('Begin ' + __name__ + ' convert2csv')

    FileKey = isolateFileKey( pFileName)
    if   FileKey == 'cables' :
         convert2csv_cables( pFileName)
    elif FileKey == 'pro' :
         convert2csv_pro( pFileName)
    elif FileKey == 'dsp' :
         convert2csv_dsp( pFileName)
    elif FileKey == 'pa'  :
         convert2csv_pa( pFileName)
    else :
        log.info('File ' + pFileName + ' - ignore')
    if os.path.exists( myname+'_'+FileKey+'.csv'):
        shutil.copy2(  myname+'_'+FileKey+'.csv', 'c://AV_PROM/prices/' + myname +'/'+ myname+'_'+FileKey+'.csv')
    


def isolateFileKey( sourceString):
    re_fname = re.compile('^.+_([^_]+)_[0-9]+\..*$', re.LOCALE | re.IGNORECASE )
    is_fname = re_fname.match(sourceString)
    if is_fname:                        # Файл соответствует шаблону имени
        key = is_fname.group(1)         # выделяю ключ из имени файла
    else:
        key = '' 
    return key.lower()
