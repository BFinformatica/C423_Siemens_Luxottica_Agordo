VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_CalibCode As String
Dim m_CalibID As String
Dim m_rsObj As Object
Dim m_Status As String
Dim m_Command As String
Dim m_StartTime As Date
Dim m_Duration As Long
Dim m_TimeRemaining As Long
Dim CounterID As Integer
Dim m_Running As Boolean


Public Property Get CurrentStatus() As String
  CurrentStatus = m_Status
End Property
Public Property Let CurrentStatus(Value As String)
  m_Status = Value
End Property
Public Property Get CurrentCommand() As String
  CurrentCommand = m_Command
End Property
Public Property Let CurrentCommand(Value As String)
  m_Command = Value
End Property
Public Property Get TimeRemaining() As Long
  TimeRemaining = m_TimeRemaining
End Property
Public Property Let TimeRemaining(Value As Long)
  m_TimeRemaining = Value
End Property
Public Property Get Duration() As Long
  Duration = m_Duration
End Property
Public Property Let Duration(Value As Long)
  m_Duration = Value
End Property
Public Property Get StartTime() As Date
  StartTime = m_StartTime
End Property
Public Property Let StartTime(Value As Date)
  m_StartTime = Value
End Property
Public Property Get CalibCode() As String
  CalibCode = m_CalibCode
End Property
Public Property Let CalibCode(Value As String)
  m_CalibCode = Value
End Property
Public Property Get CalibID() As String
  CalibID = m_CalibID
End Property
Public Property Let CalibID(Value As String)
  m_CalibID = Value
End Property
Public Property Get IsRunning() As Boolean
  IsRunning = m_Running
End Property
Public Property Let IsRunning(Value As Boolean)
  m_Running = Value
End Property





Public Property Let Counter(Value As Integer)
  CounterID = Value
End Property
Public Property Get Counter() As Integer
  Counter = CounterID
End Property



Public Property Get rsObject() As Object
  Set rsObject = m_rsObj
End Property

Private Sub Class_Initialize()
  NewDataObj m_rsObj
End Sub

Private Sub Class_Terminate()
  Set m_rsObj = Nothing
  CalibObjCount = CalibObjCount - 1
End Sub
