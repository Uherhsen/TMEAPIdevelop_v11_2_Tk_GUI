# -*- coding: utf-8 -*-
"""
Списки описаний
В случае ошибки проверка ДБ на сервисе: https://jsonlint.com/
"""

import json

def my_descrlist(jsonname):
    '''Проверяет наличие ДБ, гружает '''
    try:
        data = json.load(open(jsonname))
        return list(data[1].keys())
    except:
        print("Отсутствует или поврежден файл ДБ "+jsonname+'\nВ случае ошибки проверка ДБ на сервисе: https://jsonlint.com/'+'\n')
        
if __name__ == "__main__":
    filename = 'database.json'
    print('Не отсортированный:\n\n',my_descrlist(filename),'\n\nкол-во: ', len(my_descrlist(filename)))
    # sorted(my_descrlist(filename))