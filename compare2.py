import os,sys
import argparse
#import pyexcel as pe
import xlrd
import xlsxwriter
import types
import utl

ver = "0.0"


#lines = None   # -- Обмеження числа зчитуваних рядків (для відладки)
args = None  # -- Для збереження оргументів командного рядка глобально



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
        bg_arg  - рядок з аргументами з даними про початок області порівняння.
        Наприклад, А3 (одна колонка)  cd3 (дві колонки)
        '''
        self.sheet = sheet
        self.bg_arg = bg_arg
        self.bg = utl.get_range_bg(bg_arg) # [rov1,coli1,coli2]

        if maxlines:
            self.maxlines = int(maxlines)          # -- обмеження числа зчитуваних рядків
        else:
            self.maxlines = None
            
        bg = self.bg

        self.first_row = bg[0]  # -- індексний номер першого рядка
        self.cols = bg[1:]      # -- індексні номери задіяних колонок
        print ('cols:',self.cols)

        self.end_row = self.sheet.nrows  # -- Індекс на одиницю більший за індекс останнього рядка
        print ('end row:',self.end_row)


    def get_dict(self):
        ''' Будує словник значень для діапазона (з врахуванням обмежень).
            Ключ - пошуковий рядок.
            Значення - номер рядка.
        Повертає побудований словник.
        '''
        res = {}
        nrows = 0
        for nrow in range(self.first_row,self.end_row):
            row = self.sheet.row_values(nrow)
            # Формуємо зведене значення колонок
            nval = self.get_cols_value(row,self.cols) # -- отримує приведене значення
            #print nrow,nval
            res[nval] = nrow
            nrows += 1
            if self.maxlines and nrows > self.maxlines:
                break
                
        return res


    def get_cols_value(self,row,ncols):
        ''' Формує спільне значення з кількох кологок.
        row - список значень рядка
        ncols - номоери задіяних колонок
        '''
        res = ''
        for ncol in ncols:
            cell_value = self.norm(row[ncol])
            res += cell_value
        return res


    def norm(self,value):
        ''' Нормалізує подане значення:
        - приводить до символьного вигляду, якщо це число.
        - (мождиво знадобиться ще щось, наприклад очистка від пробілів чи приведення до одного регістру)
        Повертає нормалізоване значення
        '''
        print ('--norm value type:',type(value))
        print ('--norm value "%s"' % value)
        if type(value) == str:
            res = value
        else:
            res = str(value)
        return res


    def dump(self):
        ''' Виводить дані
        '''
        print (" -- Range input argument:",self.bg_arg)
        print ("-- bg:",self.bg)
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

    validate_args(file1,file2)
    # -- Зчитуємо обидва ексель файли
    
    
    wb1 = xlrd.open_workbook(file1)
    sheet1 = wb1.sheet_by_index(0)
    print ("--sheet1 rows:",sheet1.nrows)
    print ("--sheet1 cols:",sheet1.ncols)

    wb2 = xlrd.open_workbook(file2)
    sheet2 = wb2.sheet_by_index(0)
    print ("--sheet2 rows:",sheet2.nrows)
    print ("--sheet2 cols:",sheet2.ncols)

    # -- Створюємо дескриптори діапазонів для кожного листа

    range1 = SheetRange(sheet1,bg1,args.lines)
    range2 = SheetRange(sheet2,bg2,args.lines)

    # -- Побудова словників
    d1 = range1.get_dict()
    d2 = range2.get_dict()

    print ('--Dict1:',d1)
    print ('--Dict2:',d2)
    
    # Дивимось результат
    range1.dump()
    

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Порівнює вказані колонки із дв ох ексель файлів та видає результат у залежності від режиму порівняння')
    parser.add_argument("file_1_name",help="Перший файл")
    parser.add_argument("file_2_name",help="Другий файл")
    parser.add_argument("col_file_1",help="Початкова клітинка першого файла") 
    parser.add_argument("col_file_2",help="Початкова клітинка другого файла")  

    parser.add_argument("-l","--lines",help="Обмеження числа зчитуваних рядків вхідних файлів")
    
    args = parser.parse_args()


    print ('Args:',args)
    print('Args type:',type(args))

    # ---- *** Додати валідацію аргументів ****
        
    
    # Отримуємо назви обох файлів
    file_1 = args.file_1_name # -- Назва першого файла
    file_2 = args.file_2_name # -- Назва другого файла
                           # ---- **** Зробити можливість задати тільки колонку
                           #      **** без вказування рядка
    bg1 = args.col_file_1  # -- Початок у першому файлі (наприклад 'B1')
    bg2 = args.col_file_2  # -- Початок у другому файлі    


    
    main(file_1,file_2,bg1,bg2,args)
    






