# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 12:02:49 2020

Получение данных о электронных компонентах с сайта TME через его API
"""
import openpyxl,time
from TME_Python_API import product_import_tme


# Функция считает все ячейки первого столбца в которых что то написано,до тех пор пока не встретит пустую ячейку "None"
def number_of_articles(path):
    """Функция считает все ячейки первого столбца в которых что то написано,до тех пор пока не встретит пустую ячейку None"""
    # Открываем Эксель
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    #выясняем количество артикулов в файле эксель
    i = 1
    while sheet['A'+str(i)].value != None:
        i+=1
    wb.save(path)
    return i-1

# Функция создающая список артикулов. Принимает число артикулов и номер колонки в виде буквы (str): 'A'- первая колонка
#
def articlesList(n,column,path):
    """ Функция создающая список артикулов. Принимает число артикулов и номер колонки в виде буквы (str): 'A'- первая колонка"""
    # Открываем Эксель
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active 
    cord_in = column+str(1)
    cord_out = column+str((n)) 
    # формирование списка артикулов
    vals = [v[0].value for v in sheet[cord_in : cord_out]] #[r[0].value for r in sheet.Range(cord)]
    return vals
    wb.save(path)
#   
# Функция для вывода пинга
#   
def ping():
    """Функция для вывода пинга"""
    ping_data = product_import_tme(token, app_secret, 'Utils/Ping', params={})
    print(ping_data)
#        
# Функция проставляет оригинальные артикулы производителя, дескрипшен, ссылку на фото, ссылку на страницу продукта и вес в КГ! 
#            
def search_articles(articles_list, path, params, token, app_secret,action1, rng1=0):
    '''Поиск базовых параметров, задача найти оригинальный артикул использование экшена Products/Search'''
    print('Поиск базовых параметров')
    # Открываем Эксель
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    rng2=len(articles_list)
    params2 = params.copy()
    index = 0
    for j in range(rng1,rng2):
        global flag
        if flag != 1: # Флаг нужен для принудительного завершения цикла, меняется из графического интерфейса
            break
        else:
            index += 1
            print( "Акт 1", index, 'Поиск базовых параметров')
            params['SearchPlain'] = str(articles_list[j])
            all_data = product_import_tme(token, app_secret, action1, params)
            #print(all_data)
            try :
                print(all_data['Data']['ProductList'][0]['Symbol'],all_data['Data']['ProductList'][0]['Description'])
                all_data['Data']['ProductList'][0]['Symbol']
                sheet['B'+str(j+1)] = all_data['Data']['ProductList'][0]['Symbol']
                sheet['C'+str(j+1)] = all_data['Data']['ProductList'][0]['Description']
                if all_data['Data']['ProductList'][0]['Photo'] == "":
                    print('нет фото')
                else:
                    sheet['E'+str(j+1)] = "https://" + all_data['Data']['ProductList'][0]['Photo'][2:] # добавляем КАРТИНКИ!
                
                sheet['F'+str(j+1)] = all_data['Data']['ProductList'][0]['ProductInformationPage'][2:]
                weight = (all_data['Data']['ProductList'][0]['Weight'])
                if all_data['Data']['ProductList'][0]['WeightUnit']=='g':
                    weight = weight*0.001
                    #print('{:f}'.format(weight))
                    weight = round(weight,(('{:f}'.format(weight)).count('0'))+1) # округление
                sheet['D'+str(j+1)] = weight
                
            except IndexError: # ____________________________________________________________________________________________________В этом блоке запускаются другие экшены
                if all_data["Status"]=="OK":
                    
                    params2['SymbolList[0]'] = articles_list[j]
                    action3 = 'Products/GetProductsFiles'
                    try:
                        all_data = product_import_tme(token, app_secret, action3, params2 ) # используем другой экшн для нахождения картинок Products/GetProductsFiles 
                        
                    except:
                        print ('\nСтатус сети ',all_data["Status"],'\nАртикул "',articles_list[j],'" отсутствует на TME\n')
                        sheet['C'+str(j+1)] = 'Артикула нет на TME'
                    else:
                        
                        sheet['E'+str(j+1)] = "https:" + all_data['Data']['ProductList'][0]['Files']['PhotoList'][0]
                        sheet['B'+str(j+1)] = articles_list[j] # Копируем во-второй столбец артикул для поиска другими экшенами
                    try:
                        action0 = 'Products/GetProducts' # получаем веса по-другому экшену
                        all_data = product_import_tme(token, app_secret, action0, params2 ) # используем другой экшн для нахождения картинок Products/GetProductsFiles 
                        weight = (all_data['Data']['ProductList'][0]['Weight'])
                        if all_data['Data']['ProductList'][0]['WeightUnit']=='g':
                            weight = weight*0.001
                            weight = round(weight,(('{:f}'.format(weight)).count('0'))+1) # округление
                    except:
                        print ('\nСтатус сети ',all_data["Status"],'\nАртикул "',articles_list[j],'" нет весов\n')
                    else: # добавляем Веса 2
                        sheet['D'+str(j+1)] = weight
                else:
                    print('\nСтатус сети ',all_data["Status"])
                    time.sleep(2) 
                #continue
    wb.save(path)
    #time.sleep(0.1)
    print('\nПоиск базовых параметров завершен')
#    
# функция использует "экшен" для поиска параметров, к которому нужны оригинальные артикулы,
# запускается только после функции search_articles
#
def search_param(articles_list,path,params,token,app_secret,action2, rng1=0):
    """ Использование экшена Products/GetParameters"""
    # Открываем Эксель
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    rng2=len(articles_list)
    print('\nЦикл проставления параметров\n')
    index = 0
    for j in range(rng1,rng2):
        global flag
        if flag != 1:
            break
        else:
            index += 1
            print( "Акт 2", index, 'Цикл проставления параметров')
            if articles_list[j] != None:
                params['SymbolList[0]'] = articles_list[j]
                
                try:
                    all_data = product_import_tme(token, app_secret, action2, params) # НЕ РАБОТАЕТ если не находящийся на ТМЕ артикул, по какой то причине есть во второй колонке !!!
                except: # пропускаем принудительный выход raise
                    print(' HTTPError\n_____________________________________________________')
                    continue
                if all_data['Status'] == "OK":
                    try:
                        print(all_data['Data']['ProductList'][0]['ParameterList'][1]['ParameterName'],
                              all_data['Data']['ProductList'][0]['ParameterList'][1]['ParameterValue'])
                    except IndexError:
                        print("\nОшибка структуры ответа\n")
                        continue
                    prms={}
                    for i in all_data['Data']['ProductList'][0]['ParameterList']:
                        prms[i['ParameterName']] = i['ParameterValue']
                    manufacturer = prms.pop('Производитель','Нет ключа')
                    #print(manufacturer)
                    sheet['K'+str(j+1)] = str(manufacturer)
                    prms.pop('#Promotion','Нет ключа')
                    prms.pop('#Promotion','Нет ключа')
                    prms.pop('вес','Нет ключа')
                    sheet['G'+str(j+1)] = str(prms)
                    
                    #time.sleep(0.3)
                else:
                    print('ошибка статуса сети')
                    j+=1
                      
                    
            else:
                print('Пропуск артикула')
                j+=1
    wb.save(path)
    #time.sleep(0.1)
    print('\nЦикл проставления параметров завершен')
#
# Проставление ссылок на даташит
#
def products_files(articles_list, path,params,token,app_secret,action3, rng1=0):
    # Открываем Эксель
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    rng2=len(articles_list)
    print('\nЦикл поиска ссылок на даташит\n')
    index = 0
    for j in range(rng1,rng2):
        global flag
        if flag != 1:
            break
        else:
            index += 1
            print( "Акт 3", index, 'Цикл поиска ссылок на даташит')
            if articles_list[j] != None:
                params['SymbolList[0]'] = articles_list[j]
                try:
                    all_data = product_import_tme(token, app_secret, action3, params)
                except: # пропускаем принудительный выход raise
                    print(' HTTPError\n_____________________________________________________')
                    continue
                if all_data['Status'] == "OK":
                    try:
                        print(all_data['Data']['ProductList'][0]['Files']['DocumentList'][0]['DocumentUrl'][2:],'\n'+('_'*50)) #[0]['DocumentUrl'])
                        sheet['H'+str(j+1)] = "https://"+ str(all_data['Data']['ProductList'][0]['Files']['DocumentList'][0]['DocumentUrl'][2:])
                    except KeyError:
                        print('KeyError')
                        j+=1
                    except IndexError:
                        print('Нет PDF:\n',all_data['Data']['ProductList'][0]['Files'],'\n'+('_'*50))
                        j+=1
                else:
                    print('ошибка статуса сети')
                    break
                
                try: # Добываем Оригинальный артикул, будем считать что это артикул производителя
                    OriginalSymbol = product_import_tme(token, app_secret, 'Products/GetProducts', params)
                    sheet['L'+str(j+1)]= OriginalSymbol['Data']['ProductList'][0]['OriginalSymbol']
                except: # пропускаем принудительный выход raise
                    print("Нет оригинального наименования")
                    continue
            else:
                print('Нет артикула','\n'+('_'*50))
                j+=1
    print('\nЦикл поиска ссылок завершен')
    wb.save(path)
#       
if __name__ == '__main__': 
    flag=1
    xlsxpath = "productdata.xlsx"
    
    params={'Country' : 'RU','Language' : 'RU',}
    token = 'ac434c181917ed4e51c49a2027bfd040e9f2da0054be7'
    app_secret = '0b748f6e5d340d693703'
    action0 = 'Products/GetProducts' # request method, метод пинг Utils/Ping, Products/Search
    action1 = 'Products/Search'
    action2 = 'Products/GetParameters'
    action3 = 'Products/GetProductsFiles'
    n = number_of_articles(xlsxpath)
    
    work_articles_list_A = articlesList(n, 'A', xlsxpath)    
    search_articles(work_articles_list_A, xlsxpath,params,token, app_secret, action1)
    work_articles_list_B = articlesList(n,'B', xlsxpath)
    search_param(work_articles_list_B, xlsxpath,params,token,app_secret,action2)
    products_files(work_articles_list_B, xlsxpath,params,token,app_secret,action3)