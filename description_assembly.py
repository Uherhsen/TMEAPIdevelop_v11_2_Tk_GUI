# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 16:00:40 2020

Программа для сборки описания по информации из экселя и из шаблонов JSON БД (Работает уже с заполненным экселем)
"""
from get_product_name import product_type
from config import replace_list # Импорт переменной-списка замен
from GetDataTME_with_openpyxl import GetData as Data
import openpyxl,json,configparser
#
# Проверка в БД json ключей и добавление новых, ОБЯЗАТЕЛЬНО К ВЫПОЛНЕНИЮ
def checkKey(xlsxpath,jsonfilename):
    list_range = Data.number_of_articles(xlsxpath) # диапазон
    descr_list = Data.articlesList(list_range,"G", xlsxpath) # список параметров
    #
    '''Проверка и обновление типов продуктов в базе JSON - database.json'''
    #
    # Список типов из экселя формируются из параметров (колонка G), из слов c ключем вида "Тип ...чего-то там : ...что-то там,"(Функция product_type)
    categ_list = [product_type(d) for d in descr_list]
    # Получаем словарь значений из базы
    try:
        data = json.load(open(jsonfilename)) # FileNotFoundError:
    except FileNotFoundError:
        print('Файл базы database.json отсутствует, или неверно указанно имя файла')
    # Получаем список ключей из словаря базы json (ключи это типы продукции, - первые слова дескрипшина только уже помещенные в базу )
    data_keys = list(data[0].keys()) 
    '''подгружает новые элементы, присваивает значения None (null стандарт джейсона)'''
    for e in categ_list:
        # Добавляем данные в словарь, избегая повторений
        if e not in data_keys: data[0][e] = None
    # Звгружаем обновленный словарь в базу json
    with open(jsonfilename,'w') as file:
        json.dump(data, file, indent=2, ensure_ascii=False)
#
''' Сборка описания. 
БД database.json содержит список из двух словарей [{Тип продукта:ссылка-ключ},{ссылка-ключ:[начало, конец описания,[[это заменить, заменить на это],...[это заменить, заменить на это]]]}]: 
    первый словарь - это множество типов товаров полученное отделением параметра "тип" от списка параметров (столбец экселя G) в качестве ключей, по которым получаем значение-ссылку являющюся ключем второго словаря по которому получаешь описание, 
    второй словарь - по ключу отдает описание в виде списка [0-Начало описания,1-конец описания,[[заменяемое,заменитель],[..., ...]...[..., ...]]].
описание составляется конкатенацией: нулевой элемент словаря + параметры из экселя + первый элемент словаря, при этом происходят общие замены (некоторые символы не допускаются в описании)
затем происходят частные замены, которые указываются в индивидульном списке в описании из БД [2]-индекс элемента
Если во втором словаре при обращении по ключу первый (с нулевым индексом) элемент словаря == "MY_FUNСTION", то запускается альтернативный сценарий, который распаковывает строку-имя функции с индексом [1] в этом же словаре'''

#
# для замен, принимает массив замент типа [[a,b],[a,b],...[a,b]] и текст, в котором a будет заменено на b
def replaceAB(replList,text):
    """Для замен, принимает массив замент типа [[a,b],[a,b],...[a,b]] и текст, в котором a будет заменено на b"""
    for a,b in replList:
        text = text.replace(a,b)
    return text        
# Для альтернативного сценария: В ячейку сохранены только параметры, без описания.
def just_parameters(i):
    print("\nВ ячейку сохранены только параметры, без описания: \n\n"+ prod_param+".\n")
    sheet["I"+str(i)]= 'Параметры без описания: '+prod_param+'.'
#    wb.save(xlsxpath)    
# Для альтернативного сценария: для вставки параметров в текстовый шаблон
def textTemplate(i,link_key): 
    #print(data[1][link_key][1])
    # получаем словарь параметров из экселя, i-из цикла
    prms = eval(sheet['G'+str(i)].value)
    # получаем шаблон описания из JSONа, по ключу-ссылке к имени функции (link_key - берется из основного цикла)
    descTemplate = data[1][link_key][2]
    # получаем словарь значений по-умолчанию из JSONа, для случаев осутствия необходимых вставляемых параметров
    prmsDflt = data[1][link_key][3]
    for key in [*prmsDflt]:
        if key not in prms:
            prms[key]= prmsDflt[key]
    d=descTemplate.format(prms) 
    try:
        replaceL = data[1][link_key][4]
        replaceL=replaceL+replace_list
        for a,b in replaceL:
            d = d.replace(a,b)
        print('По текстовому шаблону:\n '+d)
        #return d
        sheet["I"+str(i)]= d #Сдесь добавлялась запись для тестов 'По текстовому шаблону: '+
    except IndexError:
        print('По текстовому шаблону: нет списка замен\n '+d)
        #return d
        for a,b in replace_list:
            d = d.replace(a,b)
        sheet["I"+str(i)]= 'По текстовому шаблону, отсутствует список замен: '+d
#    wb.save(xlsxpath)
#
#============================ ОСНОВНОЙ СЦЕНАРИЙ =========================>
#
def desc_assembly(xlsxpath,jsonfilename):
    '''Основной сценарий создания описания, производит конкатенацию строк, замены и пр.'''
    global wb
    wb = openpyxl.load_workbook(xlsxpath)#путь к файлу
    global sheet
    sheet = wb.active # лист экселя
    list_range = Data.number_of_articles(xlsxpath) # количество позиций
    # Получаем словарь значений из базы
    try:
        global data
        data = json.load(open(jsonfilename)) # FileNotFoundError:
    except FileNotFoundError:
        print('Файл базы database.json отсутствует, или неверно указанно имя файла')
    for i in range(1, list_range+1):
        print(i)
        # Отделяем категорию от дескрипшина
        name = product_type(sheet['G'+str(i)].value) # типы продукта !!!
        # Помещаем в колонку J тип продукции
        sheet['J'+str(i)] = name
        #print(name)
        global prod_param
        prod_param = sheet['G'+str(i)].value
        if prod_param != None:
            try:
                # Переменная-ссылка-ключ к словарю с описаниями
                link_key = data[0][name]
                # Проверка наличия флага "MY_FUNСTION" для включения иных сценариев
                if data[1][link_key][0] == "MY_FUNСTION":
                    # Включение функции написанной в базе
                    eval(data[1][link_key][1]) # запуск альтернативного сценария
                    continue
            except KeyError:
                sheet["I"+str(i)]= None
                print("Не указан шаблон-описание")
                continue
            else:
                # Получение словаря
                # превращаем в словарь текст
                params_dict = eval(prod_param)
                # Избавляемся от ненужного теперь ключа 'Категория продукта' 
                #params_dict.pop('Категория продукта','Нет ключа')
                # УДАЛЕНИЯ ПО КЛЮЧУ
                #
                try:
                    for pop_param in data[1][link_key][3]:
                        params_dict.pop(pop_param,'Нет ключа')
                except IndexError:
                    print('Без замен по ключу')
                
                # Преобразование словаря в текст
                strparam=[]
                # Получаем словарь
                for dict_item in list( params_dict.items()):
                    dict_item=dict_item[0]+' '+ dict_item[1] #.lower()
                    strparam.append(dict_item)
                comma = ', '
                # преобразуем словарь в текст
                prod_param = comma.join(strparam)
                # цикл общих замен  
                prod_param = replaceAB(replace_list,prod_param)
                # Сборка описания: КОНКАТЕНАЦИЯ
                try:
                    my_descr = data[1][link_key][0]+prod_param+data[1][link_key][1]
                except TypeError:
                    print('Шаблон-описание пуст')
                    continue
                except IndexError:
                    print('Ошибка в шаблоне-описании (IndexError)')
                    continue
                # ЧАСТНЫЕ ЗАМЕНЫ  (по словаою из БД джесон)
                try:
                    if data[1][link_key][2] == []:
                        sheet["I"+str(i)]= my_descr
                    else:
                        my_descr=replaceAB(data[1][link_key][2],my_descr) 
                        #for x,y in data[1][link_key][2]: # остатки после рефакторинга
                            #my_descr = my_descr.replace(x,y)
                        print('С доп. заменой:\n', my_descr)
                        sheet["I"+str(i)]= ' '+my_descr # С доп. заменой КОНКАТЕНИРУЕТ МАРКЕР В ТЕКСТЕ СООБЩАЮЩИЙ О ДОПОЛНИТЕЛЬНЫХ ЗАМЕНЫХ
                except IndexError:
                    # Если в БД джейсон нет блока с заменами к этому типу продукции, то оставляем все как есть
                    print(my_descr)
                    sheet["I"+str(i)]= my_descr
            
        else:
            print('Параметров продукта нет на TME')
            continue   
    wb.save(xlsxpath)
    print('\n\n Сборка описания завершена')

if __name__ == '__main__':
    path = "productdata.xlsx"
    filename = 'database.json'
    checkKey(path,filename)
    desc_assembly(path,filename)