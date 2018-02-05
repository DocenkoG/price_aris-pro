# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time
import aris_pro_converter
import aris_pro_downloader
import shutil
import requests, lxml.html
from aris_pro_converter    import convert2csv
import re

global log
global myname


def download( cfg ):
    global myname
    pathDwnld = './tmp'
    retCode     = False
    filename_new= cfg.get('download','filename_new')
    filename_old= cfg.get('download','filename_old')
    if cfg.has_option('download','login'):     login       = cfg.get('download','login'    )
    if cfg.has_option('download','password' ): password    = cfg.get('download','password' )
    if cfg.has_option('download','href_text'): href_text   = cfg.get('download','href_text' )
    url_download_page= cfg.get('download','url_download_page'   )
    url_base         = cfg.get('download','url_base' )
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.0; rv:14.0) Gecko/20100101 Firefox/14.0.1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Language':'ru-ru,ru;q=0.8,en-us;q=0.5,en;q=0.3',
               'Accept-Encoding':'gzip, deflate',
               'Connection':'keep-alive',
               'DNT':'1'
              }
    try:
        s = requests.Session()
        r = s.get(url_download_page,  headers = headers)
        page = lxml.html.fromstring(r.text)
        url_file = None
        for item in page.xpath('//a'):
            if item.text == href_text:
                print(item.attrib)
                url_file = item.get('href')
                print(url_file)
        if url_file == None :
            log.error('Не найден элемент %s на странице %s', href_text, url_download_page)
            return False
        r = s.get(url_base +'/'+ url_file)
        log.debug('Загрузка файла %16d bytes   --- code=%d', len(r.content), r.status_code)
        retCode = True
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    if os.path.exists( filename_new) and os.path.exists( filename_old): 
        os.remove( filename_old)
        os.rename( filename_new, filename_old)
    if os.path.exists( filename_new) :
        os.rename( filename_new, filename_old)
    f2 = open(filename_new, 'wb')                                  # Теперь записываем файл
    f2.write(r.content)
    f2.close()
    if filename_new[-4:] == '.zip':                                # Архив. Обработка не завершена
        log.debug( 'Zip-архив. Разархивируем '+ filename_new)
        work_dir = os.getcwd()
        if not os.path.exists('tmp'):   os.mkdir('tmp')                                       
        os.chdir( os.path.join( pathDwnld ))
        dir_befo_download = set(os.listdir(os.getcwd()))
        print( dir_befo_download)
        os.remove( '*.xls*')
        dir_befo_download = set(os.listdir(os.getcwd()))
        print( dir_befo_download)
        os.system('unzip -oj ' + filename_new)
        dir_afte_download = set(os.listdir(os.getcwd()))
        new_files = list( dir_afte_download.difference(dir_befo_download))
        print(new_files)

    for new_file in new_files :
        new_ext  = os.path.splitext(new_file)[-1]
        DnewFile = new_file
        new_file_date = os.path.getmtime(DnewFile)
        log.debug( 'Файл из архива ' +DnewFile + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )
        DnewPrice = DnewFile
        FoldName = 'old_' + new_file                                        # Предыдущая копия прайса, для сравнения даты
        FnewName = 'new_' + new_file                                        # Файл, с которым работает макрос
        if  (not os.path.exists( FnewName)) or new_file_date>os.path.getmtime(FnewName) : 
            log.debug( 'Предыдущего прайса нет или он устарел. Копируем новый.' )
            if os.path.exists( FoldName): os.remove( FoldName)
            if os.path.exists( FnewName): os.rename( FnewName, FoldName)
            os.rename(DnewFile, FnewName)
            #
            #       Вызов конвертации файла
            #
            retCode = convert2csv( FnewName, 'aris_pro')
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
#§======================================================                        
    return retCode
    return new_files[0]



def isolateFileKey( sourceString):
    re_fname = re.compile('^.+_([^_]+)_[0-9]+\..*$', re.IGNORECASE )
    is_fname = re_fname.match(sourceString)
    if is_fname:                        # Файл соответствует шаблону имени
        key = is_fname.group(1)         # выделяю ключ из имени файла
    else:
        key = '' 
    return key.lower()



def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days *24*60*60 
    if os.path.exists( fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    if price_datetime+qty_seconds < time.time() :
        file_age = round((time.time()-price_datetime)/24/60/60)
        log.error('Файл "'+fileName+'" устарел!  Допустимый период '+ str(qty_days)+' дней, а ему ' + str(file_age) )
        return False
    else:
        return True



def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def processing(cfgFName):
    log.info('----------------------- Processing '+cfgFName )
    cfg = config_read(cfgFName)
    #filename_out  = cfg.get('basic','filename_out')
    #filename_in= cfg.get('basic','filename_in')
    
    if cfg.has_section('download'):
        result = download(cfg)
    if is_file_fresh( filename_in, int(cfg.get('basic','срок годности'))):
        #os.system( dealerName + '_converter_xlsx.xlsm')
        convert2csv(cfg)
    folderName = os.path.basename(os.getcwd())
    #if os.path.exists( filename_out): shutil.copy2( filename_out, 'c://AV_PROM/prices/' +folderName+'/'+filename_out)
    if os.path.exists( 'python.log'): shutil.copy2( 'python.log', 'c://AV_PROM/prices/' +folderName+'/python.log')
    if os.path.exists( 'python.1'  ): shutil.copy2( 'python.log', 'c://AV_PROM/prices/' +folderName+'/python.1'  )
    


def main( dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый 
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          '+dealerName )
    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            processing(cfgFName)


if __name__ == '__main__':
    global myname
    myname = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myname)
    main( myname)
