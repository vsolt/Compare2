import os,sys
import argparse
#import pyexcel as pe
import xlrd
import xlsxwriter
import types
import utl
import result

ver = "0.1"


args = None  # -- Для збереження оргументів командного рядка глобально

def console(utf_message):
    ''' Друкує на консоль повідомлення. Якщо програма виконуєтьсяу віндовсі,
    спочатку перекодовує у CP866 (з укр-i корекцією)
    utf_message : Рядок у кодуванні utf-8
    ver-0.1
    '''
    #if os.name == 'nt':
    #    res = utf_message.replace(u'і',u'i')
    #    res = res.replace(u'І',u'I')
    #    res = res.encode('cp866')
    #else:
    #    res = utf_message
    print (utf_message)


def validate_args(file1,file2):
    ''' Перевірка наявності вказаних файлів.
    При невідповідності друкує на консоль помилку і завершує роботу.
    При успішній валідації повертає True
    '''
    for file in (file1,file2):
        if not os.path.isfile(file):
            #console_message('Помилка в назві файла %s. Такого файла не існує.' % file)
            print('Помилка в назві файла: %s. Такого файла не існує.' % file) # ***
            sys.exit()
    return True   # *** Додати валідацію колонок


class SheetRange(object):
    ''' Містить дані про діапазон клітинок ексель-файла
    '''
    def __init__(self,sheet,bg_arg,maxlines=None):
        ''' Задає координати першої точки
        bg_arg  - рядок з аргументами з даними про початок області
        порівняння. Наприклад, А3 (одна колонка)  cd3 (дві колонки)
        maxlines - обмеження числа рядків (для відладки)
        '''
        self.sheet = sheet
        self.bg_arg = bg_arg # -- 
        #print('--Range bg_arg:',bg_arg)
        self.bg = utl.get_range_bg(bg_arg) # [row_index,col_index]
        #print ('Range row,col:',self.bg)
        self.first_row = self.bg[0]  # -- індексний номер першого рядка
        
        
        if maxlines:
             self.maxindex = int(maxlines) -1    # -- Максимальний індекс зчитуваного рядка
        else:
            self.maxindex = None
            
        self.cols = self.bg[1:]      # -- індексні номери задіяних колонок
        #print ('cols:',self.cols)

        self.end_row = self.sheet.nrows  # -- Індекс на одиницю більший за індекс останнього рядка
        #print ('end row:',self.end_row)


    def get_dict(self):
        ''' Будує словник значень для діапазона (з врахуванням обмежень).
            Ключ - пошуковий рядок.
            Значення - номер рядка.
        Повертає побудований словник.
        '''
        res = {}
        count = 0 # -- лічильник для обмеження числа рядків
        for nrow in range(self.first_row,self.end_row):
            
            row = self.sheet.row_values(nrow)
            
            # Формуємо зведене значення колонок  ***
            key = self.get_cols_value(row,self.cols) # -- отримує приведене значення
            #print ('rownum,key,value:',nrow,key,nrow)
            
            # -- Якщо ключ вже існує, дописуємо номер рядка
            if not key in res:
                res[key] = [nrow]
            else:
                res[key].append(nrow)
            count += 1
            
            #if self.maxindex >= 0 and count > self.maxindex:
            if not self.maxindex is None  and count > self.maxindex:
                break
                
        return res


    def get_cols_value(self,row,ncols):
        ''' Формує спільне значення з кількох колонок.  ***
        row - список значень рядка
        ncols - номомери задіяних колонок
        '''
        res = ''
        for ncol in ncols:
            cell_value = self.norm(row[ncol])
            res += cell_value
        return res


    def norm(self,value):
        ''' Нормалізує подане значення:
        - 
        - приводить до символьного вигляду, якщо це число.
        - (мождиво знадобиться ще щось, наприклад очистка від пробілів чи приведення до одного регістру)
        Повертає нормалізоване значення
        '''
        #print ('--norm value type:',type(value))
        res = value
        if type(res) == float:
            res = int(res)
        if not type(res) == str:
            res = str(res)
        # print ('--norm res:',res,type(res))
        return res


    def dump(self):
        ''' Виводить дані
        '''
        print (" -- Command argument:",self.bg_arg)
        print ("-- First row index:",self.bg)
        print ("cols",self.cols)
        print ("--end row_i:",self.end_row)
        print ("--first row:", self.first_row)

        #print "--first:(%s,%s)" % (self.first_col,self.first_row)
        #print "--last:(%s,%s)" % (self.last_col,self.end_row)



def main(file1,file2,bg1,bg2,args):
    ''' Обробка двох файлів.
    file1,file2 : назви першого та другого файла
    bg1 : Назва колонки або координата першої клітинки для порівняння у першому файлі - наприклад A або A2
    bg2 : Те ж для другого файла
    args: Об'єкт розбору аргументів  командного рядка
   
    '''
    print("Тест друку кирилицею - українська розкладка також дивимось і")
    validate_args(file1,file2)
    # -- Зчитуємо обидва ексель файли
    
    
    wb1 = xlrd.open_workbook(file1)
    sheet1 = wb1.sheet_by_index(0)
    console("Число рядків у першому файлі:%s" % sheet1.nrows)

    wb2 = xlrd.open_workbook(file2)
    sheet2 = wb2.sheet_by_index(0)
    #rows2 = sheet2.nrows
    console("Число рядків у другому файлі:%s" % sheet2.nrows)    

    # -- Створюємо дескриптори діапазонів для кожного листа

    range1 = SheetRange(sheet1,bg1,args.lines)
    range2 = SheetRange(sheet2,bg2,args.lines)

    # -- Побудова словників з індексами
    d1 = range1.get_dict()
    d2 = range2.get_dict()

    #print ('--Dict1:',d1)    
    #print ('--Dict2:',d2)

    
    # Дивимось результат
    #range1.dump()

    # -- Пошук відсутніх у другому
    absent2 = utl.get_absent(d1,d2)  # -- Список номерів відсутніх рядків
    #print ('absent2:',absent2)

    # -- Пошук у першому
    absent1 = utl.get_absent(d2,d1)
    #print ('absent1:',absent1)

    # -- Виписка номерів рядків першого файла ключі яких є у другому файлі
    common1 = utl.get_common(d1,d2)
    #print('Common1:',common1)

    # -- Виписка номерів рядків другого файла ключі яких є у першому файлі
    common2 = utl.get_common(d2,d1)
    #print('Common2:',common2)    


    
    # -- Побудова результату
    '''
    1. Створити workbook
    2. Створити sheets
    3. Переписати заголовок із базового sheeh
    4,5,6 Виписати рядки із базового sheet

    '''
    exgen = result.ExcelGen(range1,range2)

    # -- Пишемо рядки з першого файла, ключів яких нема у другому файлі
    console (u'Відсутніх у другому файлі:%s' % len(absent2))
    exgen.add_sheet_rows(sheet1,0,absent2)

    

    # -- Пишемо рядки, виписані із другого файла
    exgen.add_sheet_rows(sheet2,1,absent1)
    console(u'Відсутніх у першму файлі:%s' % len(absent1))

    # -- Виписуємо  рядки першого файла знайдені за ключами друго файла
    console(u'Присутніх у обох файлах:%s' % len(common1))
    exgen.add_sheet_rows(sheet1,2,common1,rep=True)

    # -- Виписуємо  рядки другого файла знайдені за ключами першого файла
    #console(u'Присутніх у обох файлах:%s' % len(common2))    
    #exgen.add_sheet_rows(sheet2,3,common2,rep=True)

    # -- Виписуємо рядки першого файла спільні наявними ключами к 2му
    rep1  = utl.get_stat_repeated(d1)
    rep2  = utl.get_stat_repeated(d2)    

    console(u'Повторних у першому файлі:%s' % rep1)
    console(u'Повторних у другому файлі:%s' % rep2)    


    

    exgen.save('out.xls')

    print ("saved")


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Порівнює вказані колонки із двох ексель файлів та видає результат у залежності від режиму порівняння')
    parser.add_argument("file_1_name",help="Перший файл")
    parser.add_argument("file_2_name",help="Другий файл")
    parser.add_argument("col_file_1",help="Початкова клітинка першого файла") 
    parser.add_argument("col_file_2",help="Початкова клітинка другого файла")  

    parser.add_argument("-l","--lines",help="Обмеження числа зчитуваних рядків вхідних файлів (для тестування)")
    
    args = parser.parse_args()


    # ---- *** Додати валідацію аргументів ****
        
    
    # Отримуємо назви обох файлів
    file_1 = args.file_1_name # -- Назва першого файла
    file_2 = args.file_2_name # -- Назва другого файла
                           # ---- **** Зробити можливість задати тільки колонку
                           #      **** без вказування рядка
    bg1 = args.col_file_1  # -- Початок у першому файлі (наприклад 'B1')
    bg2 = args.col_file_2  # -- Початок у другому файлі    


    
    main(file_1,file_2,bg1,bg2,args)
    






