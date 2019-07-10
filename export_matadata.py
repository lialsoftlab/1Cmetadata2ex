import sys
import pprint

import toml
import win32com.client as com

from huepy import red, green, grey, cyan
from ruamel.yaml import YAML

# monkey patch для поддержки русского языка в идентификаторах методов иначе win32com.client.build.MakePublicAttributeName() их вырезает.
com.build.valid_identifier_chars += 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя'

config = toml.load("export.toml")

sys.stderr.write(grey("Загрузка COM connector v8.3 ... "))
sys.stderr.flush()

v83cc = com.gencache.EnsureDispatch('V83.COMConnector')

sys.stderr.write(green("OK\n"))
sys.stderr.flush()

sys.stderr.write(grey("Соединяемся с 1С ... "))
sys.stderr.flush()

db1c = v83cc.Connect(config["connection_string"])

sys.stderr.write(green("OK\n"))
sys.stderr.flush()

sys.stderr.write(grey("Получаем структуру хранения базы данных ... "))
sys.stderr.flush()

objects = db1c.NewObject("Массив")
for metaobj in config.get("metaobjects", []):
    objects.Add(metaobj)

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

    fields = sorted(fields, key=lambda x: x.get("FLD_META", '_') + x.get("FLD_1CQL", '_'))
    
    d = dict(TAB_DBMS=line.ИмяТаблицыХранения, TAB_CLASS=line.Назначение)
    if line.ИмяТаблицы: d["TAB_1CQL"] = line.ИмяТаблицы
    if line.Метаданные: d["TAB_META"] = line.Метаданные
    if fields: d["TAB_FIELDS"] = fields 
    result.append(d)
    sys.stderr.flush()


YAML().dump(sorted(result, key=lambda x: x.get("TAB_META", '_') + x.get("TAB_1CQL", '_') + x['TAB_DBMS']), sys.stdout)

