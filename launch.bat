REM needed for start 32-bit cscript/wscript on 64-bit system else not working with 32-bit COMConnector
REM or else you need to patch 1C COM connector registration in registry according with instructions here
REM https://www.codeproject.com/Tips/267554/Using-32-bit-COM-Object-from-64-bit-Application
 
%SystemRoot%\SysWOW64\cscript.exe %1 %2 %3 %4 %5 