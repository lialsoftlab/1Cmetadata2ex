WScript.Echo "Loading COM connector v8.3 ..."
Dim cc1c
Set cc1c = CreateObject("V83.ComConnector")

WScript.Echo "Connecting to DB ..." 
Dim db1c
Set db1c = cc1c.Connect("Srvr=""192.168.6.93"";Ref=""Trade_EP"";")

WScript.Echo "Getting DB storage information ..."
For Each item In db1c.GetDBStorageStructureInfo(Empty, True) 
    WScript.Echo(item.TableName)
Next