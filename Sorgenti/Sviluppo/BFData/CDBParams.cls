VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get DBType() As String
  DBType = AppDBType
End Property
Public Property Let DBType(Value As String)
  AppDBType = Value
End Property

Public Property Get Database() As String
  DBType = AppDatabase
End Property
Public Property Let Database(Value As String)
  AppDatabase = Value
End Property

Public Property Get DBUser() As String
  DBType = AppDBUser
End Property
Public Property Let DBUser(Value As String)
  AppDBUser = Value
End Property

Public Property Get DBPwd() As String
  DBType = AppDBPwd
End Property
Public Property Let DBPwd(Value As String)
  AppDBPwd = Value
End Property

Public Property Get Server() As String
  DBType = AppServer
End Property
Public Property Let Server(Value As String)
  AppServer = Value
End Property

Public Property Get DBVersion() As String
  DBVersion = AppDBVersion
End Property
Public Property Let DBVersion(Value As String)
  AppDBVersion = Value
End Property







