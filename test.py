import sys

import win32com.client as com

from huepy import red, green, grey, cyan
from ruamel.yaml import YAML

# monkey patch для поддержки русского языка в идентификаторах методов иначе win32com.client.build.MakePublicAttributeName() их вырезает.
com.build.valid_identifier_chars += 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя'

sys.stderr.write(grey("Загрузка COM connector v8.3 ... "))
sys.stderr.flush()

v83cc = com.gencache.EnsureDispatch('V83.COMConnector')

sys.stderr.write(green("OK\n"))
sys.stderr.flush()

sys.stderr.write(grey("Соединяемся с 1С ... "))
sys.stderr.flush()

db1c = v83cc.Connect("Srvr=\"192.168.6.93\";Ref=\"Trade_EP\";")

sys.stderr.write(green("OK\n"))
sys.stderr.flush()

sys.stderr.write(grey("Получаем структуру хранения базы данных ... "))
sys.stderr.flush()

objects = db1c.NewObject("Массив")
# objects.Add("Справочник.Товары")
# objects.Add("Справочник.Единицы")
# objects.Add("Документ.Инвойс")
# objects.Add("Документ.ЗаказПоставщику")

result_kvt = db1c.ПолучитьСтруктуруХраненияБазыДанных(objects, True)

sys.stderr.write(green("OK\n"))
sys.stderr.flush()

sys.stderr.write(grey(f"Кол-во объектов = [{green(result_kvt.Count())}]\n"))
sys.stderr.flush()

result = list()
for line in result_kvt:
    sys.stderr.write(grey(f"Таблица: {cyan(line.ИмяТаблицыХранения)} ... \n"))
    
    fields = list()
    for field in line.Поля:
        d = dict(FLD_DBMS=field.ИмяПоляХранения)
        if field.ИмяПоля: d["FLD_1CQL"] = field.ИмяПоля
        if field.Метаданные: d["FLD_META"] = field.Метаданные
        fields.append(d)
        sys.stderr.write(grey(f"\tПоле: {cyan(field.ИмяПоляХранения)}\n"))

    fields = sorted(fields, key=lambda x: x.get("FLD_META", '.') + x.get("FLD_1CQL", '.'))
    
    result.append(dict(TAB_DBMS=line.ИмяТаблицыХранения, TAB_1CQL=line.ИмяТаблицы, TAB_META=line.Метаданные, TAB_CLASS=line.Назначение, TAB_FIELDS=fields))
    sys.stderr.flush()


YAML().dump(sorted(result, key=lambda x: x['TAB_DBMS']), sys.stdout)

