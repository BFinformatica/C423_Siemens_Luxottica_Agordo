VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_LookUpKey As String
Dim m_Station As String

Public Property Let LookUpKey(Value As String)
  m_LookUpKey = Value
End Property
Public Property Get LookUpKey() As String
  LookUpKey = m_LookUpKey
End Property

Public Property Let Station(Value As String)
  m_Station = Value
End Property
Public Property Get Station() As String
  Station = m_Station
End Property


Public Sub WriteLog(strData As String, Optional bNoCrLf As Boolean)
  Dim Oggi As Date
  Dim m_LogObj As Object

  On Error GoTo err
  NewDataObj m_LogObj
  m_LogObj.m_Tabella = "wds_Log"
  
  With m_LogObj
    .AddNew
      Oggi = Now
      .SetValue "Log_Date", DateSerial(Year(Oggi), Month(Oggi), Day(Oggi))
      .SetValue "Log_Hour", Format$(Now, "hh.nn.ss")
      .SetValue "Log_String", strData
      .SetValue "Log_LookupKey", m_LookUpKey
      .SetValue "Log_Station", m_Station
    .Save
  End With
  
err:
  Set m_LogObj = Nothing
End Sub

