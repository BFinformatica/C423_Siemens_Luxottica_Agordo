VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTryArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  uCallBackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private nid As NOTIFYICONDATA
Private m_Form As Form
Private m_OriginalState As Integer

Public Sub IconAdd(FormObj As Form, Title As String)
  Set m_Form = FormObj
  With nid
      .cbSize = Len(nid)
      .hwnd = FormObj.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = FormObj.Icon
      .szTip = Title & vbNullChar
  End With
  m_OriginalState = m_Form.WindowState
  Shell_NotifyIcon NIM_ADD, nid
  
End Sub
      
Public Sub RestoreWindow()
  m_Form.WindowState = m_OriginalState
  SetForegroundWindow m_Form.hwnd
  m_Form.Show
End Sub

Public Sub IconDelete()
  Shell_NotifyIcon NIM_DELETE, nid
End Sub




