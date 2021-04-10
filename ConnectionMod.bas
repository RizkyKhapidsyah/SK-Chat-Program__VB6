Attribute VB_Name = "ConnectionMod"
Option Explicit
Public glousername As String
Public adoconn As New ADODB.Connection
'Enter the servername of your network and give the path of the chat database
'servername = In place of pdc or your local network server
'path of the database= In place of Doctorsoft\chat.mdb create a directory and load the chat.mdb database in it.
Function condb()
adoconn.Open "DBQ=\\PDC\DoctorSoft\Chat.mdb;DefaultDir=\\PDC\DoctorSoft;Driver={Microsoft Access Driver (*.mdb)};DriverId=26;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=18;PageTimeout=15;SafeTransactions=0;Threads=5;UID=admin;UserCommitSync=Yes;"
End Function

Function conclose()
adoconn.Close
End Function



