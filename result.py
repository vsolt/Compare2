# coding:  utf-8

''' Оформлення результату порівняння файлів.
'''
import xlwt

class ExcelGen(object):
    ''' Генератор результуючого ексель-файла
    Виводить знайдені дані у файл.
    '''
    def __init__(self,range1,range2):
        ''' Ініціалізація генератора.
        range1,range2 - об`єкти SheetRange
        Створює workbook та порожні сторінки із заголовками.
        Всі потрібні дані бере із range
        '''
        self.wb = xlwt.Workbook()
        self.sheet1 = self.wb.add_sheet(u'Відсутні у 2')
        self.sheet2 = self.wb.add_sheet(u'Відсутні у 1')
        #self.sheet3 = self.wb.add_sheet(u'Присутні у 1 та 2')
        self.sheet3 = self.wb.add_sheet(u'Спільні у 1')
        self.sheet4 = self.wb.add_sheet(u'Спільні у 2')
        

        hr1 = self.add_sheet_header(self.sheet1,range1)
        hr2 = self.add_sheet_header(self.sheet2,range2)
        hr3 = self.add_sheet_header(self.sheet1,range1)
        hr4 = self.add_sheet_header(self.sheet2,range2)
        
        self.next_row1 = hr1 + 1
        self.next_row2 = hr2 + 1
        self.next_row3 = hr3 + 1
        self.next_row4 = hr4 + 1
                
        self.sheets = []
        self.sheets.append(self.sheet1)
        self.sheets.append(self.sheet2)
        self.sheets.append(self.sheet3)
        self.sheets.append(self.sheet4)



        self.row_cnts = []
        self.row_cnts.append(self.next_row1)
        self.row_cnts.append(self.next_row2)
        self.row_cnts.append(self.next_row3)
        self.row_cnts.append(self.next_row4)

    def add_sheet_header(self,sheet,s_range):
        ''' Виписує заголовок у заданий sheet
        range - SheetRange з потрібними даними
        Повертає індекс останнього записаного рядка
        '''
        base_sheet = s_range.sheet
        header_nrows = s_range.first_row
        #print "--header nrows:",header_nrows,type(header_nrows)
        res = 0
        if header_nrows:
            # -- Виписуємо рядки із заголовком
            for nrow in range(header_nrows):
                self.copy_row(nrow,base_sheet,sheet)
            res = nrow
        return res


    def copy_row(self,rowi,base_sheet,sheet):
        ''' Копіює рядок з base_sheet у sheet
        rowi   - індекс рядка
        Копіюється у ту ж позицію на sheet
        '''
        row = base_sheet.row_values(rowi)
        for coli in range(len(row)):
            sheet.write(rowi,coli,row[coli])

    def add_sheet_rows(self,base_sheet,sheet_i,row_lists,rep=False):
        ''' Переписує у вказаний sheet рядки із base_sheet.
        base_sheet : звідки виписувати
        sheet_i - індекс цільового sheet (у який писати)
        row_lists - список списків індексів рядків із base_sheet які виписати у sheet_i
                  Наприклад [[0],[5,7],...]
        rep : True коли потрібно виводити рядки з додатковими індексами
        Міняє дані лічильника рядків відповідного sheet
        '''

        # Для задіяння додаткових рядків виписуємо вкдені індекси у один список
        if rep:  # -- Друк дадаткових записів для індекса
            irows = []
            for rowlist in row_lists:
                for row in rowlist:
                    irows.append([row])
        else:
            irows = row_lists         #  ---- *****
            
        print ("--add rows to sheet:",sheet_i,len(row_lists))
        print ('--n_sheets:',len(self.sheets))
        print ('--row_lists',row_lists)
        print ('irows:',irows)
        
        sheet = self.sheets[sheet_i]      # -- Куди писати
        row_cnt = self.row_cnts[sheet_i]  # -- У який рядок писати
        
        for irow in irows:
            self.write_sheet_row(irow,base_sheet,sheet,row_cnt)
            row_cnt += 1

    def write_sheet_row(self,row_list,base_sheet,sheet,row_cnt):
        ''' Додає рядок
        row_list : список з одним або кількома індексами рядка із base_sheet
        base_sheet : звідки писати
        row_cnt : у який рядок писати
        '''
        row_i = row_list[0]
        print('write sheet row:',row_i)
        row = base_sheet.row_values(row_i)
        for coli in range(len(row)):
            #cell = row[coli]
            sheet.write(row_cnt,coli,row[coli])




    def save(self,filename):
        ''' Зберігає workbook під вказаною назвою.
        '''
        self.wb.save(filename)
