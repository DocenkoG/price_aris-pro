# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time
import shutil
import requests
#from aris_pro_converter    import convert2csv
from aris_pro_converter_cables import convert2csv_cables 
from aris_pro_converter_pro    import convert2csv_pro 
from aris_pro_converter_dsp    import convert2csv_dsp 
from aris_pro_converter_pa     import convert2csv_pa 
import re
#from unrar.cffi import rarfile

global log
global myname


def convert2csv( pFileName   # file for convertation
               , myname      # organisatiom name (dir name)
               ) :
    global log
    global SheetName
    global FilenameIn
    global FilenameOut
    make_loger()
    log.debug('Begin  convert2csv '+ pFileName)

    FileKey = isolateFileKey( pFileName)
    #cfgFName = "aris_pro_cfg_"+ FileKey+".cfg"
    #if os.path.exists(cfgFName):
    #    cfg = config_read(cfgFName)
    #    if  is_file_fresh( pFileName, int(cfg.get('basic','срок годности'))):
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
   

def download( cfg ):
    global myname
    is_download_success = False
    filename_new= cfg.get('download','filename_new')
    filename_old= cfg.get('download','filename_old')
    if cfg.has_option('download','login'):     login       = cfg.get('download','login'    )
    if cfg.has_option('download','password' ): password    = cfg.get('download','password' )
    if cfg.has_option('download','href_text'): href_text   = cfg.get('download','href_text' )
    #url_download_page= cfg.get('download','url_download_page'   )
    url_file = cfg.get('download','url_file')
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.0; rv:14.0) Gecko/20100101 Firefox/14.0.1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Language':'ru-ru,ru;q=0.8,en-us;q=0.5,en;q=0.3',
               'Accept-Encoding':'gzip, deflate',
               'Connection':'keep-alive',
               'DNT':'1'
              }
    try:
        s = requests.Session()
        '''
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
            is_download_success = False
        '''
        r = s.get(url_file)
        log.debug('Загрузка файла %16d bytes   --- code=%d', len(r.content), r.status_code)
        is_download_success = True
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    if is_download_success :
        if os.path.exists( filename_new) and os.path.exists( filename_old): 
            os.remove( filename_old)
            os.rename( filename_new, filename_old)
        if os.path.exists( filename_new) :
            os.rename( filename_new, filename_old)
        f2 = open(filename_new, 'wb')                                  # Теперь записываем файл
        f2.write(r.content)
        f2.close()
    else:
        if not is_file_fresh( filename_new, int(cfg.get('download','срок годности'))):
            return False
            
    if filename_new[-4:] in ('.zip', '.rar'):                          # Архив. Обработка не завершена
        log.debug( 'Архив, разархивируем '+ filename_new)
        work_dir = os.getcwd()
        if not os.path.exists('tmp'):
            os.mkdir('tmp')                                       
        os.chdir( os.path.join( 'tmp' ))
        for f in os.listdir("."):
            if f.endswith(".xls"):
                os.remove(f)
        dir_befo_download = set(os.listdir("."))
        if filename_new[-4:] == '.zip':
            os.system('unzip -oj ' + os.path.join('..', filename_new))
        elif filename_new[-4:] == '.rar':
            os.system('"c:\\Program Files\\unrar\\UnRAR.exe" e -y ' + os.path.join('..', filename_new) + ' .\\')
        dir_afte_download = set(os.listdir("."))
        new_files = list( dir_afte_download.difference(dir_befo_download))
        print(new_files)
        
        os.chdir( '..' )
        for new_file in new_files :
            FoldName = 'old_' + new_file                                        # Предыдущая копия прайса, для сравнения даты
            FnewName = 'new_' + new_file                                        # Файл, с которым работает макрос
            if  os.path.exists( FnewName) : 
                if os.path.exists( FoldName): os.remove( FoldName)
                os.rename( FnewName, FoldName)
            print('====== 1 =====', os.path.join('tmp', new_file), FnewName)
            shutil.copy(            os.path.join('tmp', new_file), FnewName)
            print('====== 2 =====', os.path.join('tmp', new_file), FnewName)

    for new_file in new_files :
        #
        #       Вызов конвертации файла
        #
        retCode = convert2csv( 'new_' + new_file, 'aris_pro')
    return retCode



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
