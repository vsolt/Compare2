# coding:  utf-8

import re

cstr = 'abcdefghijklmnopqrstuvwxyz'
cdict = {}
for i in range(len(cstr)):
    cdict[cstr[i]] = i


def get_col_num(col):
    ''' Конвертує символьну назву колоноки у число.
    col - символьна назва колонки. Наприклад A або F
    Повертає число, що відповідає номеру колонки (a->0 b->...)
    '''

    lcol = col.lower()
    res = cdict.get(lcol,None)
    return res


def get_range_bg(str):
    ''' Конвертує символьні координати початку діапазону обробки.
    Наприклад 'A1' конвертує у (0,0)
    Повертає список [row_index,col_1_index,...,col_n_index] або None
    '''

    # -- Потрібно відділити букви від цифр
    chars,num = split_str(str)
    # -- Послідовність букв перетворити в послідовність номерів
    res = []
    res.append(num-1)
    for char in chars:
        res.append(get_col_num(char))
    return res


def split_str(str):
    ''' Розділяє рядок на символьну та числову частину.
    Наприклад, 'ab12' --> ('ab',12)
    str - рядок dble 'a1' або 'cd12'
    Повертає кортеж із двома значеннями-символьним та числовим або None
    '''
    tpl = '(\D+)(\d+)'  # -- Група нечислових символів далі група числових символів
    m = re.search(tpl,str)
    if not m:
        return None
    rstr = m.group(1)
    rnum = int(m.group(2))

    return (rstr,rnum)

def get_absent(dict1,dict2):
    ''' Повертає список з номерами рядків першого словника, ключі яких відсутні у другому словнику.
    Вхідні словники у форматі {value:[[rownum,]]}
    '''
    res = []
    for (key,rownums) in dict1.items():
        if not key in dict2:
            res.append(rownums)
    return res


def get_common(dict1,dict2):
    ''' Повертає список з номерами рядків першого словника, для яких є той же ключ у другому словнику.
    dict1 : словник з якого виписуються номери рядків
    dict2 : словник з ключами для яких виписувати рядки з першого словника
    Обидва вхідні словники у форматі {key:[[rownum]]}
    '''

    res = []
    for key,rownums in dict1.items():
    
        # -- дивимось чи є запис з таким ключем
        if key in dict2:
            res.append(rownums)
    return res











