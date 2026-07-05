import argparse
import pyexcel as pe
import xlsxwriter


ver = "0.0"

if __name__ == '__main__':

    #resultator = Resultator("result.csv")
    #writer = XlsWriter('out.xlsx')
    
    parser = argparse.ArgumentParser(description='Порівнює вказані колонки із двох ексель файлів та видає результат у залежності від режиму порівняння')
    parser.add_argument("input_file_1",help="Перший файл")
    parser.add_argument("col_file_1",help="Колонка першого файла")     
    parser.add_argument("input_file_2",help="Другий файл")
    parser.add_argument("col_file_2",help="Колонка другого файла")     
    # arser.add_argument("-l","--lines",help="Обмеження числа зчитуваних рядків вхідого файла")
    args = parser.parse_args()


    print ('Args:',args)
    print('Args type:',type(args))

    # Читаємо обидва файли
    #records_1 = pe.get_records(file_name=args.input_file_1)
    records_1 = pe.get_records(file_name='f_1.xls')
    #records_2 = pe.get_records(file_name=args.input_file_2)

    print ('file1 records:',len(records_1))
    #print ('file2 records:',len(records_2))
