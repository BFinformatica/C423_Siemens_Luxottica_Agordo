Attribute VB_Name = "AttimoMdl"
Option Explicit
Public AppServer As String
Public AppDatabase As String
Public AppDBUser As String
Public AppDBPwd As String
Public AppDBType As String
'Public AppParams As CParams
Public m_CurStation As String
Public m_CurDir  As String

Public AppServerMin As String
Public AppDatabaseMin As String
Public AppDBUserMin As String
Public AppDBPwdMin As String
Public AppDBTypeMin As String



Public Sub NewDataObj(rs As Object)
  On Error GoTo Err_CreateRsObject
  
  Set rs = CreateObject("AttimoFwk.CData")
  With rs
    Select Case AppDBType
      Case "ODBC"
        .SetDBType .Conn_ODBC
      Case "UDL"
        .UDLFile = App.Path & "\bflab7.udl"
        .SetDBType .Conn_UDL
      Case "ACCESS"
        .SetDBType .Conn_Jet
      Case "SQL"
        .SetDBType .Conn_SQL
        '.SetDBVersion AppDBVersion
        
      Case "ORACLE"
        .SetDBType .Conn_Oracle
      
      Case Else '"MYSQL"
        .SetDBType .Conn_MYSQL
    End Select
    .SetServer AppServer
    .SetDatabase AppDatabase, AppDBUser, AppDBPwd
    .SetMessages False
    .Err_Activate = True
    '.Err_Activate = False
  End With
  Exit Sub
Err_CreateRsObject:
  MsgBox (Err.Description)
End Sub


Public Sub NewDataObjMin(rs As Object)
  
  On Error GoTo Err_CreateRsObject
  
  Set rs = CreateObject("AttimoFwk.CData")
  With rs
    Select Case AppDBTypeMin
      Case "ODBC"
        .SetDBType .Conn_ODBC
      Case "UDL"
        .UDLFile = App.Path & "\bflab7.udl"
        .SetDBType .Conn_UDL
      Case "ACCESS"
        .SetDBType .Conn_Jet
      Case "SQL"
        .SetDBType .Conn_SQL
        '.SetDBVersion AppDBVersion
        
      Case "ORACLE"
        .SetDBType .Conn_Oracle
      
      Case Else '"MYSQL"
        .SetDBType .Conn_MYSQL
    End Select
    .SetServer AppServerMin
    .SetDatabase AppDatabaseMin, AppDBUserMin, AppDBPwdMin
    .SetMessages False
    .Err_Activate = True
  End With
  Exit Sub
Err_CreateRsObject:
  MsgBox (Err.Description)
End Sub
Public Sub NewParamObj(rs As Object)
On Error GoTo Err_CreateRsObject
  
  Set rs = CreateObject("AttimoFwk.CData")
  With rs
    .SetDBType .Conn_Jet
    .SetDatabase App.Path & "\bflab7.mdb"
  End With
  Exit Sub
Err_CreateRsObject:
  MsgBox (Err.Description)
End Sub

Public Sub NewWindasObj(rs As Object)
'On Error GoTo Err_CreateRsObject
'
'  Set rs = CreateObject("AttimoFwk.CData")
'  With rs
'    .SetDBType .Conn_Jet
'    .SetDatabase gsDirLavoro & "\windas03.mdb"
'  End With
'  Exit Sub
'Err_CreateRsObject:
'  MsgBox (Err.Description)
End Sub



Sub GetConnectionParams()
  Dim param As Object
  Dim Crypt As Object
  Dim ErrCount As Integer
  Dim AccessObj As Object
  Dim ErrPos As Integer
  Dim AppReadFromIni As Boolean
  
  'On Error GoTo Err_Load
  
  AppReadFromIni = True
  If (Dir$(App.Path & "\bflab7.mdb") <> "") Then
    Set AccessObj = CreateObject("AttimoFwk.CData")
    With AccessObj
RetryForProviderErr:
      .Err_Activate = True
      .SetMessages False
      Select Case ErrCount
        Case 0
          .SetDBType .Conn_Jet
          .SetDatabase App.Path & "\bflab7.mdb", "admin", ""
        Case Else
          AppReadFromIni = True
      End Select
      AppReadFromIni = False
      If (Not AppReadFromIni) Then
        AppReadFromIni = True
        .SelectionFast ("SELECT * FROM Connections WHERE DEFAULT = -1")
        If (Not .IsEOF) Then
          AppServer = .GetValue("SERVER")
          AppDatabase = .GetValue("DATABASE")
          AppDBType = .GetValue("DBTYPE")
          'AppDBVersion = .GetValue("CNMODE")
          AppDBUser = .GetValue("USER")
          AppDBPwd = .GetValue("PASSWORD")
          
          Set Crypt = CreateObject("AttimoFwk.CCrypt")
          AppDBUser = Crypt.Decrypt(AppDBUser)
          AppDBPwd = Crypt.Decrypt(AppDBPwd)
          Set Crypt = Nothing
          AppReadFromIni = False
        
        
        '***** modifica per database minuto *****
        .SelectionFast ("SELECT * FROM Connections WHERE DEFAULT = 0")
        If (Not .IsEOF) Then
          AppServerMin = .GetValue("SERVER")
          AppDatabaseMin = .GetValue("DATABASE")
          AppDBTypeMin = .GetValue("DBTYPE")
          'AppDBVersion = .GetValue("CNMODE")
          AppDBUserMin = .GetValue("USER")
          AppDBPwdMin = .GetValue("PASSWORD")
          
          Set Crypt = CreateObject("AttimoFwk.CCrypt")
          AppDBUserMin = Crypt.Decrypt(AppDBUserMin)
          AppDBPwdMin = Crypt.Decrypt(AppDBPwdMin)
          Set Crypt = Nothing
          'AppReadFromIni = False
          
        Else
            AppServerMin = AppServer
            AppDatabaseMin = AppDatabase
            AppDBTypeMin = AppDBType
            AppDBUserMin = AppDBUser
            AppDBPwdMin = AppDBPwd
        End If
        '*************
        
        End If
      End If
    End With
    Set AccessObj = Nothing
  
  End If
  
  If (AppReadFromIni) Then
    Set param = CreateObject("AttimoFwk.CParam")
    
    With param
      .ParamFile = App.Path & "\bflab7.ini"
      'If (Not .Exist) Then DlgConnection.Show vbModal
      AppServer = .GetStringIni("SERVER")
      AppDatabase = .GetStringIni("DATABASE")
      AppDBType = .GetStringIni("DBTYPE")
      'AppDBVersion = .GetStringIni("AUTHENTICATION")
      
      AppDBUser = .GetStringIni("USER")
      AppDBPwd = .GetStringIni("PASSWORD")
      Set Crypt = CreateObject("AttimoFwk.CCrypt")
      AppDBUser = Crypt.Decrypt(AppDBUser)
      AppDBPwd = Crypt.Decrypt(AppDBPwd)
      
      Set Crypt = Nothing
    End With
    
    Set param = Nothing
  End If

End Sub
