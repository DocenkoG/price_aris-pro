# -*- coding: UTF-8 -*-
import os
import logging
import logging.config
import time
import shutil
import re
from   aris_pro_converter import convert2csv
 


def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')




def download( myname ):
    global log
    pathDwnld = './tmp'
    pathPython2 = 'c:/Python27/python.exe'
    make_loger()
    retCode = []
    log.debug( 'Begin '   + __name__ + '  downLoader' )
    fUnitName = os.path.join( myname +'_unittest.py')
    if  not os.path.exists(fUnitName):
        log.debug( 'Отсутствует юниттест для загрузки прайса ' + fUnitName)
    else:
        dir_befo_download = set(os.listdir(pathDwnld))
        os.system( pathPython2 + ' ' + fUnitName)                                       # Вызов unittest'a
        dir_afte_download = set(os.listdir(pathDwnld))
        new_files = list( dir_afte_download.difference(dir_befo_download))
        if len(new_files) == 1 :   
            new_file = new_files[0]                                                     # загружен ровно один файл. 
            new_ext  = os.path.splitext(new_file)[-1]
            new_name = os.path.splitext(new_file)[0]
            DnewFile = os.path.join( pathDwnld,new_file)
            new_file_date = os.path.getmtime(DnewFile)
            log.info( 'Скачанный файл ' +DnewFile + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )
            if new_ext == '.zip':                                                       # Архив. Обработка не завершена
                log.debug( 'Zip-архив. Разархивируем.')
                work_dir = os.getcwd()                                                  
                os.chdir( os.path.join( pathDwnld ))
                dir_befo_download = set(os.listdir(os.getcwd()))
                os.system('unzip -oj ' + new_name)
                os.remove(new_name + new_ext)   
                dir_afte_download = set(os.listdir(os.getcwd()))
                new_files = list( dir_afte_download.difference(dir_befo_download))
                os.chdir(work_dir)
                for new_file in new_files :
                    new_ext  = os.path.splitext(new_file)[-1]
                    DnewFile = os.path.join( pathDwnld,new_file)
                    new_file_date = os.path.getmtime(DnewFile)
                    log.debug( 'Файл из архива ' +DnewFile + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )
                    DnewPrice = DnewFile
                    FoldName = 'old_' + new_file                                        # Предыдущая копия прайса, для сравнения даты
                    FnewName = 'new_' + new_file                                        # Файл, с которым работает макрос
                    if  (not os.path.exists( FnewName)) or new_file_date>os.path.getmtime(FnewName) : 
                        log.debug( 'Предыдущего прайса нет или он устарел. Копируем новый.' )
                        if os.path.exists( FoldName): os.remove( FoldName)
                        if os.path.exists( FnewName): os.rename( FnewName, FoldName)
                        shutil.copy2(DnewFile, FnewName)
                        #
                        #       Вызов конвертации файла
                        #
                        retCode = convert2csv( FnewName, myname)
                    else:
                        log.debug( 'Предыдущий прайс не старый, копироавать не надо.' )
                        FileKey = isolateFileKey( DnewFile)
                        f1 = myname+'_'+FileKey+'.csv'
                        f2 ='c:\\AV_PROM\\prices\\'+myname+'\\'+ f1
                        if os.path.exists(f1) : os.utime(f1, None)
                        if os.path.exists(f2) : os.utime(f2, None)
                    # Убрать скачанные файлы
                    if  os.path.exists(DnewFile):  
                        os.remove(DnewFile)
                        
        elif len(new_files) == 0 :        
            log.debug( 'Не удалось скачать файл прайса ')
        else:
            log.debug( 'Скачалось несколько файлов. Надо разбираться ...')

    return retCode



def isolateFileKey( sourceString):
    re_fname = re.compile('^.+_([^_]+)_[0-9]+\..*$', re.LOCALE | re.IGNORECASE )
    is_fname = re_fname.match(sourceString)
    if is_fname:                        # Файл соответствует шаблону имени
        key = is_fname.group(1)         # выделяю ключ из имени файла
    else:
        key = '' 
    return key.lower()
