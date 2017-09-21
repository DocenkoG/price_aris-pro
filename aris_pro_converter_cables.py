# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import io
import sys
import configparser
import time
# import openpyxl                   # Для .xlsx
import xlrd                         # для .xls


def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')

def getCell(  row       # номер строки
            , col       # номер колонки 
            , isDigit   # Признак, числовое ли значение нужно из этого поля
            , sheet     #  лист XLS
            ):
    '''
    Функция возвращает значение xls-ячейки в виде строки.    
    Для цифровых ячеек делается предварительное преобразование 
    в число (пустые и нечисловые значения преобразуются в "0")
    '''
    ccc = sheet.cell(row, col)
    cellType  = ccc.ctype
    cellValue = ccc.value
    if (isDigit == 'Y' or isDigit == True) : 
        if (cellValue == '') : 
            ss = '0'
        elif (cellType in (2,3)) :                  # numeric
            if int(cellValue) == cellValue:
                ss = str(int(cellValue))
            else :
                ss = str(cellValue)
        else :
            ss = str(int(cellValue))
    else :
        ss = str(cellValue)
    return ss



def convert2csv_cables( pFileName):
    global log
    global SheetName
    #global FilenameIn
    global FilenameOut
    global out_columns_names
    global out_columns_j
    global in_columns_j
    global colGrp
    global colSGrp
    global GrpFonti
    global BrandFonti
    global SubGrpFonti
    global HeaderFonti
    global RegularFonti
    global HeaderFontSize
    global RegularFontSize
    global SubGrpBackgroundColor
    global GrpBackgroundColor
    global strHeader
    global SubGrpFontSize
    global GrpFontSize
    make_loger()
    log.debug('Begin ' + __name__ + ' convert2csv')

    # Прочитать конфигурацию из файла
    ff = config_read( 'aris_pro_cfg_cables' )
    log.debug('Открываю файл '+ pFileName)
#   book = openpyxl.load_workbook(filename = pFileName, read_only=False, keep_vba=False, data_only=False)
    book = xlrd.open_workbook( pFileName.encode('cp1251'), formatting_info=True)
#   book = xlrd.open_workbook( os.path.join( mydir, pFileName.encode('cp1251')), formatting_info=True)
    
    log.debug('Устанавливаю страницу ' + SheetName )
    sh = book.sheet_by_name( SheetName )                     # xls
#   sh = book[SheetName]                                     # xlsx   
       
 
    ssss = []
    line_qty = 0
    log.debug('На странице %d строк' % sh.nrows)
                                                             # цикл по строкам файла
    for i in range( 0, sh.nrows) :
        line_qty += 1
        xfx = sh.cell_xf_index(i, colGrp-1)
        xf  = book.xf_list[xfx]
        bgcx  = xf.background.pattern_colour_index
        fonti = xf.font_index
        try:
            ccc = sh.cell(i, colGrp-1)
            if ccc.value == None :
                print (i, colGrp, 'Пусто!!!')
                continue
            '''                                        # Атрибуты шрифта для настройки конфига
            font = book.font_list[fonti]
            print( '---------------------- Строка', i, '-----------------------' )
            print( 'Строка', i, sh.cell(i, 1).value)
            print( 'background_colour_index=',bgcx)
            print( 'fonti=', fonti)
            print( 'bold=', font.bold)
            print( 'weight=', font.weight)
            print( 'height=', font.height)
            print( 'italic=', font.italic)
            print( 'colour_index=', font.colour_index )
            print( 'name=', font.name)
            
            continue
            '''

            if (fonti == GrpFonti):                                      # Группа 
                grpName = quoted(sh.cell(i,colGrp-1).value)   
                subGrpName = ''
                brand = 'Klotz'
                print('группа', grpName)
#            elif SubGrpFontSize == ccc.font.size :                      # Подгруппа
#                subGrpName = quoted(sh.cell(row=i,column=colSGrp).value)
#            if True == ccc.font.bold :                                  # Заголовок таблицы
#                print('ddd')
#                pass
            elif ('' == sh.cell(i, in_columns_j['цена']-1).value) :      # Пустая строка
                print('xxxx')
            else :                                                       # Информационная строка
                sss = []                                                 # формируемая строка для вывода в файл
                for outColName in out_columns_names :
                    if outColName in out_columns_j :
                        if outColName in ('закупка','продажа','цена') :
                            ss = getCell(i, out_columns_j[outColName]-1, 'Y', sh) 
                        else:
                            ss = quoted( getCell(i, out_columns_j[outColName]-1, 'N', sh))
                    else : 
                        # вычисляемое поле
                        if   outColName == 'бренд' :
                            ss = brand
                        elif outColName == 'наименование' :
                            s1 = getCell(i, in_columns_j['модель']-1,   'N', sh)
                            s2 = getCell(i, in_columns_j['описание']-1, 'N', sh)
                            ss = quoted( brand + ' ' + s1 + ' ' + s2)
                        elif outColName == 'закупка' :
                            s1 = getCell(i, in_columns_j['цена']-1,     'Y', sh)
                            ss = str( float(s1)*0.75)
                        else :
                            log.debug('Не определено вычисляемое поле: <' + outColName + '>' )
                    sss.append(ss)
    
                sss.append(brand)
                sss.append(grpName)
                ssss.append(','.join(sss))
            #else :
            #    log.debug('Нераспознана строка: <' + sh.cell(i, in_columns_j['модель']-1).value + '>' )
        except Exception as e:
            log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) )
            raise e

    log.debug('------------  Обработка строк завершена. ------------')
    
    f2 = open( FilenameOut, 'w', encoding='cp1251')
    f2.write(strHeader  + ',\n')
    data = ',\n'.join(ssss) +','
    dddd = data.encode(encoding='cp1251', errors='replace')
    data = dddd.decode(encoding='cp1251')
    f2.write(data)
    f2.close()



def config_read( myname ):
    global log
    global SheetName
    #global FilenameIn
    global FilenameOut
    global out_columns_names
    global out_columns_j
    global in_columns_j
    global colGrp
    global colSGrp
    global GrpFonti
    global SubGrpFonti
    global BrandFonti
    global HeaderFonti
    global HeaderFontSize
    global RegularFonti
    global RegularFontSize
    global SubGrpBackgroundColor
    global GrpBackgroundColor
    global strHeader
    global SubGrpFontSize
    global GrpFontSize

    cfgFName = myname + '.cfg'
    log.debug('Begin config_read ' + cfgFName )
    
    config = configparser.ConfigParser()
    if os.path.exists(cfgFName):     config.read( cfgFName)
    else : log.debug('Не найден файл конфигурации.')

    # в разделе [cols_in] находится список интересующих нас колонок и номера столбцов исходного файла
    in_columns_names = config.options('cols_in')
    in_columns_j = {}
    for vName in in_columns_names :
        if ('' != config.get('cols_in', vName)) :
            in_columns_j[vName] = config.getint('cols_in', vName) 
    
    # По разделу [cols_out] формируем перечень выводимых колонок и строку заголовка результирующего CSV файла
    temp_list = config.options('cols_out')
    temp_list.sort()

    out_columns_names = []
    for vName in temp_list :
        if ('' != config.get('cols_out', vName)) :
            out_columns_names.append(vName)
    
    out_columns_j = {}
    for vName in out_columns_names :
        tName = config.get('cols_out', vName)
        if  tName in in_columns_j :
            out_columns_j[vName] = in_columns_j[tName]
    print('-----------------------------------')
    for vName in out_columns_j :
        print(vName, '\t', out_columns_j[vName])    
    print('-----------------------------------')
    strHeader = ','.join(out_columns_names)           +',бренд,группа'  # +',подгруппа'
    print('HEAD =', strHeader)

    # считываем имена файлов и имя листа
    #FilenameIn   = config.get('input','Filename_in' )
    SheetName    = config.get('input','SheetName'   )      
    FilenameOut  = config.get('input','Filename_out')
    print('SHEET=', SheetName)
    
    # считываем признаки группы и подгруппы
    if ('' != config.get('grp_properties',  'группа')) :
        colGrp               = config.getint('grp_properties',     'группа')
    if ('' != config.get('grp_properties',  'подгруппа')) :
        colSGrp              = config.getint('grp_properties',  'подгруппа')
    if ('' != config.get('grp_properties',  'GrpFonti')) :
        GrpFonti             = config.getint('grp_properties',   'GrpFonti')
    if ('' != config.get('grp_properties',  'SubGrpFonti')) :
        SubGrpFonti          = config.getint('grp_properties','SubGrpFonti')
    if ('' != config.get('grp_properties',  'BrandFonti')) :
        BrandFonti           = config.getint('grp_properties', 'BrandFonti')
    if ('' != config.get('grp_properties',  'HeaderFonti')) :
        HeaderFonti          = config.getint('grp_properties','HeaderFonti')
    if ('' != config.get('grp_properties',  'RegularFonti')) :
        RegularFonti         = config.getint('grp_properties','RegularFonti')
    if ('' != config.get('grp_properties',  'HeaderFontSize')) :
        HeaderFontSize       = config.getint('grp_properties','HeaderFontSize')
    if ('' != config.get('grp_properties',  'RegularFontSize')) :
        RegularFontSize      = config.getint('grp_properties','RegularFontSize')
    if ('' != config.get('grp_properties',  'SubGrpFontSize')): 
        SubGrpFontSize       = config.getint('grp_properties','SubGrpFontSize')
    if ('' != config.get('grp_properties',  'GrpFontSize')) :
        GrpFontSize          = config.getint('grp_properties',   'GrpFontSize')
    if ('' != config.get('grp_properties',  'SubGrpBackgroundColor')) :
        SubGrpBackgroundColor= config.getint('grp_properties','SubGrpBackgroundColor')
    if ('' != config.get('grp_properties',  'GrpBackgroundColor')) :
        GrpBackgroundColor   = config.getint('grp_properties',   'GrpBackgroundColor')
    subgrpfontbold           = config.get('grp_properties','subgrpfontbold')
    grpfontbold              = config.get('grp_properties',   'grpfontbold')
    return 



def quoted(sss):
    if ((',' in sss) or ('"' in sss) or ('\n' in sss))  and not(sss[0]=='"' and sss[-1]=='"') :
        sss = '"'+sss.replace('"','""')+'"'
    return sss
