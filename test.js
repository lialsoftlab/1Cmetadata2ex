WScript.echo("Загрузка COM connector v8.3 ...");
var cc1c = new ActiveXObject('V83.ComConnector');
WScript.echo("Соединяемся с 1С ...");
var db1c = cc1c.Connect("Srvr=\"192.168.6.93\";Ref=\"Trade_EP\";");

WScript.echo("Получаем структуру хранения базы данных ...");

//var result = db1c.GetDBStorageStructureInfo(undefined, true);
//var result = db1c.ПолучитьСтруктуруХраненияБазыДанных(undefined, true);

var objects = db1c.NewObject("Массив");
objects.Add("Справочник.Товары");
var result = db1c.ПолучитьСтруктуруХраненияБазыДанных(objects, true);
WScript.Echo("Кол-во объектов = ", result.Количество());

for (i = 0; i < result.Count(); i++) {
    line = result.Get(i);
    WScript.echo(line.ИмяТаблицы, " | ", line.ИмяТаблицыХранения, " | ", line.Метаданные, " | ", line.Назначение);
//    WScript.echo(result.Get(i).Get(0), " | ", result.Get(i).Get(1));    
}
