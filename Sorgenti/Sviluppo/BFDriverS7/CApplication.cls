VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim CalibObj As CCalibration
Dim lInstance As Integer
Dim lCalibID As String
Dim lMeasureCode As String
Dim lStart As Date
Dim lEnd As Date
Dim lStation As String


Public Property Get DB_Type() As String
  DB_Type = AppDBType
End Property
Public Property Let DB_Type(value As String)
  AppDBType = value
End Property

Public Property Get DB_Database() As String
  DB_Database = AppDatabase
End Property
Public Property Let DB_Database(value As String)
  AppDatabase = value
End Property

Public Property Get DB_User() As String
  DB_User = AppDBUser
End Property
Public Property Let DB_User(value As String)
  AppDBUser = value
End Property

Public Property Get DB_Pwd() As String
  DB_Pwd = AppDBPwd
End Property
Public Property Let DB_Pwd(value As String)
  AppDBPwd = value
End Property

Public Property Get DB_Server() As String
  DB_Server = AppServer
End Property
Public Property Let DB_Server(value As String)
  AppServer = value
End Property

Public Property Get DB_Version() As String
  DB_Version = AppDBVersion
End Property
Public Property Let DB_Version(value As String)
  AppDBVersion = value
End Property

Public Property Get Station() As String
  Station = lStation
End Property
Public Property Let Station(value As String)
  lStation = value
End Property


Public Function ExecuteCalibration(CalibCode As String) As Integer
  Dim Compensation As Boolean
  Dim rsObject As Object
  On Error GoTo ErrManager
  Set CalibObj = New CCalibration
  
   
  NewDataObj rsObject
  CalibObj.CalibCode = CalibCode
  
  With rsObject
    .SelectionFast ("SELECT * FROM wcs_EngineHeaders WHERE ech_Code = " & .parSQLStr(CalibCode))
    If (Not .IsEOF) Then
      lMeasureCode = .GetValue("ech_Measure")
      Compensation = .GetValue("ech_Compensation")
      lStart = Now
      lCalibID = CalibCode & Format$(lStart, "yyyymmddhhnnss")
      .ExecuteSQL "UPDATE wcs_EngineHeaders SET ech_ExecuteMode = 0 WHERE ech_Code = " & .parSQLStr(CalibCode)

      .SelectionFast ("SELECT * FROM wcs_EngineRows WHERE ecr_HeaderCode = " & .parSQLStr(CalibCode) & "ORDER BY ecr_RowID ")
      If (Not .IsEOF) Then
        CalibObj.StartTime = Now
        CalibObj.Duration = .GetValue("ecr_Duration")
        Instances = Instances + 1
        lInstance = Instances - 1
        ReDim Preserve frmArray(Instances)
        
        Set frmArray(lInstance) = New Form1
        
        Set frmArray(lInstance).CalibObj = CalibObj
        Set frmArray(lInstance).rsObject = rsObject
        frmArray(lInstance).Show
        frmArray(lInstance).Timer1.Enabled = True
        frmArray(lInstance).Hide
      End If
    Else
      CalibObj.CurrentStatus = "@@@"
    End If
  End With
  Set rsObject = Nothing
  
  Exit Function
ErrManager:
  Err.Raise Err.Number, Err.Description
End Function

Public Sub EndCalibration()
  Dim rsAvg As Object
  Dim rsIns As Object
  Dim strSQL As String
  On Error Resume Next
  
  lEnd = Now

  NewDataObj rsIns
  With rsIns
    strSQL = "INSERT INTO wcs_calibration(csd_StationCode,csd_calibcode,CsD_CalibID,csd_measure,csd_counter,csd_action,csd_date,csd_value,csd_Hour)" & _
             "SELECT " & .parSQLStr(lStation) & ",cid_calibcode,CID_CalibID,cid_measure,cid_counter,cid_action,cid_date,AVG(cid_value),max(cid_Hour) FROM wcs_instant " & _
             "WHERE CID_CalibID = " & .parSQLStr(lCalibID) & " " & _
             "GROUP BY cid_measure,cid_counter,cid_action,cid_calibcode,cid_date"
    
    .ExecuteSQL strSQL
  End With
  Set rsIns = Nothing
  Set CalibObj = Nothing
  Unload frmArray(lInstance)
  Set frmArray(lInstance) = Nothing
End Sub


Public Property Get CurrentStatus() As String
  CurrentStatus = CalibObj.CurrentStatus
End Property
Public Property Get CurrentCommand() As String
  CurrentCommand = CalibObj.CurrentCommand
End Property

Public Property Get TimeRemaining() As Long
  TimeRemaining = CalibObj.TimeRemaining
End Property

Public Property Get Duration() As Long
  Duration = CalibObj.Duration
End Property

Public Property Get StartTime() As Date
  StartTime = CalibObj.StartTime
End Property

Public Property Get EndTime() As Date
  EndTime = lEnd
End Property

Public Property Get CalibCode() As String
  CalibCode = CalibObj.CalibCode
End Property

Public Property Get ID() As String
  ID = lCalibID
End Property

Public Property Get Measure() As String
  Measure = lMeasureCode
End Property

Public Property Get CounterID() As Integer
  CounterID = CalibObj.Counter
End Property

