import sys
import pprint

import toml
import win32com.client as com

from huepy import red, green, grey, cyan
from ruamel.yaml import YAML

# monkey patch для поддержки русского языка в идентификаторах методов иначе win32com.client.build.MakePublicAttributeName() их вырезает.
com.build.valid_identifier_chars += 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя'


class Application:

    def __init__(self):
        self.load_config()

    def load_config(self):
        try:
            sys.stderr.write(grey("Загрузка конфигурации ... "))
            sys.stderr.flush()

            self.config = toml.load("export.toml")

            sys.stderr.write(green("OK\n"))
            sys.stderr.flush()
        except Exception as ex:
            sys.stderr.write(red(f"{str(ex)}\n\n"))
            sys.stderr.flush()
            raise ex

    def connect_with_1C(self):
        try:
            sys.stderr.write(grey("Загрузка COM connector v8.3 ... "))
            sys.stderr.flush()
        
            self.v83cc = com.gencache.EnsureDispatch('V83.COMConnector')
            
            sys.stderr.write(green("OK\n"))
            sys.stderr.flush()

            sys.stderr.write(grey("Соединяемся с 1С ... "))
            sys.stderr.flush()

            self.db1c = self.v83cc.Connect(self.config["connection_string"])

            sys.stderr.write(green("OK\n"))
            sys.stderr.flush()
        except Exception as ex:
            sys.stderr.write(red(f"{str(ex)}\n\n"))
            sys.stderr.flush()
            raise ex
    
    def extract_metadata(self):
        try:
            sys.stderr.write(grey("Получаем структуру хранения базы данных ... "))
            sys.stderr.flush()

            objects = self.db1c.NewObject("Массив")
            for metaobj in self.config.get("metaobjects", []):
                objects.Add(metaobj)

            self.result_kvt = self.db1c.ПолучитьСтруктуруХраненияБазыДанных(objects, True)

            sys.stderr.write(green("OK\n"))
            sys.stderr.flush()
        except Exception as ex:
            sys.stderr.write(red(f"{str(ex)}\n\n"))
            sys.stderr.flush()
            raise ex

    def parse_metadata(self):
        result = list()
        for line in self.result_kvt:
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

        return sorted(result, key=lambda x: x.get("TAB_META", '_') + x.get("TAB_1CQL", '_') + x['TAB_DBMS'])

    def run(self):
        self.connect_with_1C()
        self.extract_metadata()
        return self.parse_metadata()


if __name__ == "__main__":
    app = Application()
    YAML().dump(app.run(), sys.stdout)

