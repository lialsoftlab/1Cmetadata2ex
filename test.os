﻿V83CC = Новый COMОбъект("V83.COMConnector");
Сообщить(V83CC);

СтрокаСоединенияСБД = "Srvr=""192.168.6.93"";Ref=""Trade_EP_Today_COPY"";";
Сообщить(СтрокаСоединенияСБД);

БД = V83CC.Connect(СтрокаСоединенияСБД);
Для Каждого СтрокаТЗ Из БД.ПолучитьСтруктуруХраненияБазыДанных(Неопределено, Истина) Цикл
	Сообщить(СтрокаТЗ.ИмяТаблицы);
КонецЦикла
	
