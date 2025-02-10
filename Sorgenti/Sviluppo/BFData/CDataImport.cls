VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsIns As Object
'Dim rsIns2 As Object
Dim m_ValParam As String
Dim m_ValData As String
Dim m_ValOra As String
Dim m_ValValore As String
Dim m_ValStatus As String
Dim m_ValValidita As String
Dim m_ValUniMis As String
Dim m_ValParameter As String
Dim m_ValNr As String
Dim m_ValMax As String
Dim m_ValMin As String
Dim m_ValSpan As String
Dim m_ValZero As String
Dim m_ValSpanT As String
Dim m_ValZeroT As String
Dim m_ValM As String
Dim m_ValQ As String
Dim m_ValError As String
Dim strSQL As String
Dim m_StationCode As String
Dim m_SystemCode As String
Dim ErrObj As CLog
Dim ParamObj As CDBParams
Dim m_ErrMsg As String
Dim m_Table As String


Private Sub Class_Initialize()
  On Error Resume Next
  'NewDataObj rsIns
'  NewDataObj2 rsIns2
  Set ParamObj = New CDBParams
  Set ErrObj = New CLog
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set rsIns = Nothing
'  Set rsIns2 = Nothing
  Set ErrObj = Nothing
End Sub

Public Property Get DBParams() As Object
  Set DBParams = ParamObj
End Property


Public Property Get ErrMsg() As String
  ErrMsg = m_ErrMsg
End Property


Public Function Import(Filename As String, StationCode As String, Optional DataType As Integer = 0) As Boolean
'  If (AppRs Is Nothing) Then
'    NewDataObj AppRs
'    AppRs.SelectionFast "SELECT * FROM wds_auto WHERE 1=0"
'    NewDataObj rsIns
'  End If
'
'  Select Case DataType
'    Case 0
'      Import = ImportData(Filename, StationCode, Automatic)
'    Case 1
'      Import = ImportData(Filename, StationCode, Elaborated)
'    Case 2
'      Import = ImportAlarm(Filename, StationCode)
'    Case 3
'      Import = ImportCalibration(Filename, StationCode)
'    Case 4
'      Import = ImportStatus(Filename, StationCode)
'    Case 5
'      Import = ImportMinMax(Filename, StationCode)
'    Case 6
'      Import = ImportData(Filename, StationCode, HalfHour)
'    Case Else
'      Import = False
'  End Select
'End Function




End Function


Public Function ImportAlarm(Filename As String, StationCode As String) As Integer
  Dim ReaderObj As Object
  Dim CurErr As Integer
  Dim result As Long
  Dim i As Integer
  
  On Error GoTo Err_AlarmImport
  m_StationCode = StationCode
  
  ErrObj.Station = StationCode
  ErrObj.LookUpKey = "IMPORT_ALARM"
  ErrObj.WriteLog ("Importazione allarmi")
  
  m_SystemCode = ""
  With rsIns
    If (.SelectionFast("SELECT NT_SYSTEM FROM WDS_NETWORK WHERE NT_STATION = " & .ParSQLStr(StationCode) & "AND (NT_PARAM = '' OR NT_PARAM = '*' OR NT_PARAM IS NULL ) AND (NT_Component = '' OR NT_Component = '*' OR NT_Component = '' OR NT_Component IS NULL )")) Then
      m_SystemCode = .GetValue("NT_SYSTEM")
    End If
    .m_Tabella = "wds_alarm"
  End With
  
  
'  If (AppServer2 <> "") Then
'    With rsIns2
'      If (.SelectionFast("SELECT NT_SYSTEM FROM WDS_NETWORK WHERE NT_STATION = " & .ParSQLStr(StationCode) & "AND (NT_PARAM = '' OR NT_PARAM = '*' OR NT_PARAM IS NULL ) AND (NT_Component = '' OR NT_Component = '*' OR NT_Component = '' OR NT_Component IS NULL )")) Then
'        m_SystemCode = .GetValue("NT_SYSTEM")
'      End If
'      .m_Tabella = "wds_alarm"
'    End With
'
'  End If
  
  Set ReaderObj = CreateObject("AttimoFwk.CFileReader")
  With ReaderObj
    If (.OpenFile(Filename)) Then
      .SetDataMode .TabDelimited
      'Salta 2 righe
      For i = 0 To 2
        .ReadLine
      Next
      CurErr = 1
      Do While .ReadLine
        m_ValData = Trim$(Replace(ReaderObj.GetColumnValue(0), "'", ""))
        m_ValData = Mid$(m_ValData, 7, 4) & Mid$(m_ValData, 4, 2) & Left$(m_ValData, 2)
        
        m_ValOra = Replace(Left$(Trim$(ReaderObj.GetColumnValue(1)), 8), ":", ".")
        If (Len(m_ValOra) > 8) Then
          m_ValOra = Left$(m_ValOra, 8)
        End If
        
        m_ValParam = ReaderObj.GetColumnValue(2)
        m_ValValore = ReaderObj.GetColumnValue(3)
        result = result + InsertUpdateAlarm()
      Loop
    End If
    .CloseFile
    
    ImportAlarm = IIf(result > 0, True, False)
  End With
  Set ReaderObj = Nothing
  Exit Function
Err_AlarmImport:
  If (CurErr > 0) Then
    ErrObj.WriteLog ("Errore in import allarmi - " & err.Description)
    Resume Next
  Else
    ErrObj.WriteLog ("Errore in import allarmi - " & err.Description)
    Set ReaderObj = Nothing
  End If

End Function

Private Function InsertUpdateAlarm() As Integer
  On Error GoTo Err_InsUpdAlrm
  
  If (m_ValValore <> "NULL") Then
    With rsIns
      .Log_Activate = False
      .Err_Activate = True
      .SetMessages False
      .m_SelectWhere = "AL_System = " & .ParSQLStr(m_SystemCode) & " AND " & _
                       "AL_Station = " & .ParSQLStr(m_StationCode) & " AND " & _
                       "AL_Date = " & .ParSQLStr(m_ValData) & " AND " & _
                       "AL_Hour = " & .ParSQLStr(m_ValOra) & " AND " & _
                       "AL_Description = " & .ParSQLStr(m_ValParam)
                       
      .SelectionFast ("SELECT AL_Hour FROM " & .m_Tabella & " WHERE " & .m_SelectWhere)
      If (.IsEOF) Then
        strSQL = " INSERT INTO " & .m_Tabella & "(AL_System,AL_Station,AL_Date,AL_Hour,AL_Description,AL_StatusDesc) " & _
                 " VALUES (" & _
                 .ParSQLStr(m_SystemCode) & "," & _
                 .ParSQLStr(m_StationCode) & "," & _
                 .ParSQLStr(m_ValData) & "," & _
                 .ParSQLStr(m_ValOra) & "," & _
                 .ParSQLStr(m_ValParam) & "," & _
                 .ParSQLStr(m_ValValore) & ")"
        .ExecuteSQL strSQL
      Else
        strSQL = "UPDATE " & .m_Tabella & " SET " & _
                 " AL_Description = " & .ParSQLStr(m_ValParam) & "," & _
                 " AL_StatusDesc = " & .ParSQLStr(m_ValValore) & _
                 " WHERE " & .m_SelectWhere
        .ExecuteSQL strSQL
      End If
    End With
    
'    If (AppServer2 <> "") Then
'      With rsIns2
'        .Log_Activate = False
'        .Err_Activate = True
'        .SetMessages False
'        .m_SelectWhere = "AL_System = " & .ParSQLStr(m_SystemCode) & " AND " & _
'                         "AL_Station = " & .ParSQLStr(m_StationCode) & " AND " & _
'                         "AL_Date = " & .ParSQLStr(m_ValData) & " AND " & _
'                         "AL_Hour = " & .ParSQLStr(m_ValOra) & " AND " & _
'                         "AL_Description = " & .ParSQLStr(m_ValParam)
'
'        .SelectionFast ("SELECT AL_Hour FROM " & .m_Tabella & " WHERE " & .m_SelectWhere)
'        If (.IsEOF) Then
'          strSQL = " INSERT INTO " & .m_Tabella & "(AL_System,AL_Station,AL_Date,AL_Hour,AL_Description,AL_StatusDesc) " & _
'                   " VALUES (" & _
'                   .ParSQLStr(m_SystemCode) & "," & _
'                   .ParSQLStr(m_StationCode) & "," & _
'                   .ParSQLStr(m_ValData) & "," & _
'                   .ParSQLStr(m_ValOra) & "," & _
'                   .ParSQLStr(m_ValParam) & "," & _
'                   .ParSQLStr(m_ValValore) & ")"
'          .ExecuteSQL strSQL
'        Else
'          strSQL = "UPDATE " & .m_Tabella & " SET " & _
'                   " AL_Description = " & .ParSQLStr(m_ValParam) & "," & _
'                   " AL_StatusDesc = " & .ParSQLStr(m_ValValore) & _
'                   " WHERE " & .m_SelectWhere
'          .ExecuteSQL strSQL
'        End If
'      End With
'
'    End If
  End If
  InsertUpdateAlarm = 1
  Exit Function
Err_InsUpdAlrm:
  Resume Next
  'err.Raise err.Number, "InsertUpdateAlarm : ", err.Description & vbCrLf & strSQL
  'On Error GoTo 0

End Function


Public Function ImportCalibration(Filename As String, StationCode As String) As Integer
  Dim ReaderObj As Object
  Dim CurErr As Integer
  Dim result As Long
  Dim i As Integer
  
  On Error GoTo Err_CalibImport
  
  m_StationCode = StationCode
    
  ErrObj.Station = StationCode
  ErrObj.LookUpKey = "IMPORT_CALIBRATION"
  ErrObj.WriteLog ("Import calibrazioni")
  
  m_SystemCode = ""
  With rsIns
    If (.SelectionFast("SELECT NT_SYSTEM FROM WDS_NETWORK WHERE NT_STATION = " & .ParSQLStr(StationCode) & "AND (NT_PARAM = '' OR NT_PARAM = '*' OR NT_PARAM IS NULL) AND (NT_Component = '' OR NT_Component = '*' OR NT_Component IS NULL) ")) Then
      m_SystemCode = .GetValue("NT_SYSTEM")
    End If
    .m_Tabella = "wds_calibration"
  End With
  
'  If (AppServer2 <> "") Then
'    With rsIns2
'      If (.SelectionFast("SELECT NT_SYSTEM FROM WDS_NETWORK WHERE NT_STATION = " & .ParSQLStr(StationCode) & "AND (NT_PARAM = '' OR NT_PARAM = '*' OR NT_PARAM IS NULL) AND (NT_Component = '' OR NT_Component = '*' OR NT_Component IS NULL) ")) Then
'        m_SystemCode = .GetValue("NT_SYSTEM")
'      End If
'      .m_Tabella = "wds_calibration"
'    End With
'  End If
  
  Set ReaderObj = CreateObject("AttimoFwk.CFileReader")
  With ReaderObj
    If (.OpenFile(Filename)) Then
      .SetDataMode .TabDelimited
      'Salta 3 righe
      For i = 0 To 2
        .ReadLine
      Next
         
      CurErr = 1
      Do While .ReadLine
        m_ValData = Trim$(Replace(ReaderObj.GetColumnValue(0), "'", ""))
        m_ValData = Mid$(m_ValData, 7, 4) & Mid$(m_ValData, 4, 2) & Left$(m_ValData, 2)
        
        m_ValOra = Replace(Left$(Trim$(ReaderObj.GetColumnValue(1)), 8), ":", ".")
        If (Len(m_ValOra) > 5) Then
          m_ValOra = Left$(m_ValOra, 5)
        End If
        
        'Codice Parametro
        m_ValParam = Replace(ReaderObj.GetColumnValue(2), ",", ".")
        
        'Zero
        m_ValZero = Replace(ReaderObj.GetColumnValue(4), ",", ".")
        'Span
        m_ValSpan = Replace(ReaderObj.GetColumnValue(5), ",", ".")
        'Zero Teorico
        m_ValZeroT = Replace(ReaderObj.GetColumnValue(6), ",", ".")
        'Span Teorico
        m_ValSpanT = Replace(ReaderObj.GetColumnValue(7), ",", ".")
        
        'Status
        m_ValStatus = Replace(ReaderObj.GetColumnValue(10), ",", ".")
        
        'Descrizione
        m_ValValore = Replace(ReaderObj.GetColumnValue(11), ",", ".")
        
        'Valore Coeff.
        m_ValM = Replace(ReaderObj.GetColumnValue(8), ",", ".")
        'Valore Intercetta
        m_ValQ = Replace(ReaderObj.GetColumnValue(9), ",", ".")
        'Errori
        m_ValError = Replace(ReaderObj.GetColumnValue(11), ",", ".")
        
        result = result + InsertUpdateCalibra
      Loop
    End If
    .CloseFile
  End With
  Set ReaderObj = Nothing
  Exit Function

Err_CalibImport:
  If (CurErr > 0) Then
    ErrObj.WriteLog ("Errore in import allarmi - " & err.Description)
    Resume Next
  Else
    ErrObj.WriteLog ("Errore in import allarmi - " & err.Description)
    Set ReaderObj = Nothing
  End If
End Function

Private Function InsertUpdateCalibra() As Integer
  On Error GoTo Err_InsUpdCalib

  If (m_ValValore <> "NULL") Then
    With rsIns
      .Err_Activate = True
      .SetMessages False
      .Log_Activate = False
      
      .m_SelectWhere = "CL_System = " & .ParSQLStr(m_SystemCode) & " AND " & _
                       "CL_Station = " & .ParSQLStr(m_StationCode) & " AND " & _
                       "CL_Date = " & .ParSQLStr(m_ValData) & " AND " & _
                       "CL_Hour = " & .ParSQLStr(m_ValOra) & " AND " & _
                       "CL_Parameter = " & .ParSQLStr(m_ValParam)
      .SelectionFast ("SELECT CL_HOUR FROM " & .m_Tabella & " WHERE " & .m_SelectWhere)
      If (.IsEOF) Then
        strSQL = " INSERT INTO " & .m_Tabella & "(CL_System,CL_Station,CL_Date,CL_Hour,CL_Parameter,CL_Zero,CL_Span1,CL_TZero,CL_TSpan1,CL_Status,CL_Description,CL_m,CL_q,CL_Error) " & _
                 " VALUES (" & _
                 .ParSQLStr(m_SystemCode) & "," & _
                 .ParSQLStr(m_StationCode) & "," & _
                 .ParSQLStr(m_ValData) & "," & _
                 .ParSQLStr(m_ValOra) & "," & _
                 .ParSQLStr(m_ValParam) & "," & _
                 .ParSQLStr(m_ValZero) & "," & _
                 .ParSQLStr(m_ValSpan) & "," & _
                 .ParSQLStr(m_ValZeroT) & "," & _
                 .ParSQLStr(m_ValSpanT) & "," & _
                 .ParSQLStr(m_ValStatus) & "," & _
                 .ParSQLStr(m_ValValore) & "," & _
                 .ParSQLStr(m_ValM) & "," & _
                 .ParSQLStr(m_ValQ) & "," & _
                 .ParSQLStr(m_ValError) & ")"
        .ExecuteSQL strSQL
      Else
        strSQL = "UPDATE " & .m_Tabella & " SET " & _
                 " CL_Description = " & .ParSQLStr(m_ValValore) & "," & _
                 " CL_Status = " & .ParSQLStr(m_ValStatus) & "," & _
                 " CL_Zero = " & .ParSQLStr(m_ValZero) & "," & _
                 " CL_Span1 = " & .ParSQLStr(m_ValSpan) & "," & _
                 " CL_TZero = " & .ParSQLStr(m_ValZeroT) & "," & _
                 " CL_TSpan1 = " & .ParSQLStr(m_ValSpanT) & "," & _
                 " CL_m = " & .ParSQLStr(m_ValM) & "," & _
                 " CL_q = " & .ParSQLStr(m_ValQ) & "," & _
                 " CL_Error = " & .ParSQLStr(m_ValError) & _
                 " WHERE " & .m_SelectWhere
        .ExecuteSQL strSQL
      End If
    End With
    
'    If (AppServer2 <> "") Then
'      With rsIns2
'        .Err_Activate = True
'        .SetMessages False
'        .Log_Activate = False
'
'        .m_SelectWhere = "CL_System = " & .ParSQLStr(m_SystemCode) & " AND " & _
'                         "CL_Station = " & .ParSQLStr(m_StationCode) & " AND " & _
'                         "CL_Date = " & .ParSQLStr(m_ValData) & " AND " & _
'                         "CL_Hour = " & .ParSQLStr(m_ValOra) & " AND " & _
'                         "CL_Parameter = " & .ParSQLStr(m_ValParam)
'        .SelectionFast ("SELECT CL_HOUR FROM " & .m_Tabella & " WHERE " & .m_SelectWhere)
'        If (.IsEOF) Then
'          strSQL = " INSERT INTO " & .m_Tabella & "(CL_System,CL_Station,CL_Date,CL_Hour,CL_Parameter,CL_Zero,CL_Span1,CL_TZero,CL_TSpan1,CL_Status,CL_Description,CL_m,CL_q,CL_Error) " & _
'                   " VALUES (" & _
'                   .ParSQLStr(m_SystemCode) & "," & _
'                   .ParSQLStr(m_StationCode) & "," & _
'                   .ParSQLStr(m_ValData) & "," & _
'                   .ParSQLStr(m_ValOra) & "," & _
'                   .ParSQLStr(m_ValParam) & "," & _
'                   .ParSQLStr(m_ValZero) & "," & _
'                   .ParSQLStr(m_ValSpan) & "," & _
'                   .ParSQLStr(m_ValZeroT) & "," & _
'                   .ParSQLStr(m_ValSpanT) & "," & _
'                   .ParSQLStr(m_ValStatus) & "," & _
'                   .ParSQLStr(m_ValValore) & "," & _
'                   .ParSQLStr(m_ValM) & "," & _
'                   .ParSQLStr(m_ValQ) & "," & _
'                   .ParSQLStr(m_ValError) & ")"
'          .ExecuteSQL strSQL
'        Else
'          strSQL = "UPDATE " & .m_Tabella & " SET " & _
'                   " CL_Description = " & .ParSQLStr(m_ValValore) & "," & _
'                   " CL_Status = " & .ParSQLStr(m_ValStatus) & "," & _
'                   " CL_Zero = " & .ParSQLStr(m_ValZero) & "," & _
'                   " CL_Span1 = " & .ParSQLStr(m_ValSpan) & "," & _
'                   " CL_TZero = " & .ParSQLStr(m_ValZeroT) & "," & _
'                   " CL_TSpan1 = " & .ParSQLStr(m_ValSpanT) & "," & _
'                   " CL_m = " & .ParSQLStr(m_ValM) & "," & _
'                   " CL_q = " & .ParSQLStr(m_ValQ) & "," & _
'                   " CL_Error = " & .ParSQLStr(m_ValError) & _
'                   " WHERE " & .m_SelectWhere
'          .ExecuteSQL strSQL
'        End If
'      End With
'
'    End If
  End If
  InsertUpdateCalibra = 1
  
  Exit Function
Err_InsUpdCalib:
  Resume Next
  'err.Raise err.Number, "InsertUpdateCalib : ", err.Description & vbCrLf & strSQL
  'On Error GoTo 0
End Function


'1.3.13
Public Function ImportStatus(Filename As String, StationCode As String) As Integer
  Dim ReaderObj As Object
  Dim Intesta() As String
  Dim UniMis() As String
  Dim param() As String
  Dim i As Integer
  Dim CurErr As Integer
  Dim result As Long
  Dim anno As String
  Dim Table As String
  On Error GoTo Err_StatusImport
  ErrObj.Station = StationCode
  ErrObj.LookUpKey = "STATUS IMPORT"
  ErrObj.WriteLog "Import status"
  DoEvents
  m_StationCode = StationCode
  
  rsIns.m_Tabella = TABELLA
  Table = TABELLA
   
  Set ReaderObj = CreateObject("AttimoFwk.CFileReader")
  With ReaderObj
  
    If (.OpenFile(Filename)) Then
      On Error GoTo 0
      On Error GoTo ErrUpdateSt
      .SetDataMode .TabDelimited
      'Riga da saltare
      .ReadLine
      DoEvents
      'Riga intestazione contente i codici dei parametri
      .ReadLine
      ReDim Intesta(.ColumnCount)
      For i = 0 To .ColumnCount
        Intesta(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      .ReadLine
      ReDim UniMis(.ColumnCount)
      For i = 0 To .ColumnCount
        UniMis(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      .ReadLine
      ReDim param(.ColumnCount)
      For i = 0 To .ColumnCount
        param(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      'Riga da saltare
      '.ReadLine
         
      CurErr = 1
      Do While .ReadLine
        DoEvents
        m_ValOra = Replace(Trim$(.GetColumnValue(0)), ",", ".")
        'If (IsNumeric(m_ValOra)) Then
          'BFLAB Esporta le ore nel formato 1-24
          'BFDesk le gestisce nel formato 0-23
          'm_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")
          
        m_ValOra = Left$(Replace(m_ValOra, ":", "."), 5)
        'If (DataType <> Automatic) Then _
        '  m_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")

        m_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")
        m_ValData = Trim$(Replace(ReaderObj.GetColumnValue(1), "'", ""))
        m_ValData = Mid$(m_ValData, 7, 4) & Mid$(m_ValData, 4, 2) & Left$(m_ValData, 2)
        
'        #If MULTIYEAR = 1 Then
'          '21/02/2006
'          Anno = Left$(m_ValData, 4)
'          With rsIns
'            If (.Table_Exist(Table & Anno)) Then
'              .m_Tabella = Table & Anno
'            Else
'              .m_Tabella = Table
'            End If
'          End With
'
'          With rsIns2
'            If (.Table_Exist(Table & Anno)) Then
'              .m_Tabella = Table & Anno
'            Else
'              .m_Tabella = Table
'            End If
'          End With
'
'          '21/02/2006
'        #End If
        
        For i = 2 To .ColumnCount
          DoEvents
          If (Intesta(i) <> "") Then
            m_ValParam = Intesta(i)
            m_ValValore = Trim$(ReaderObj.GetColumnValue(i))
            'm_ValValore = Replace(IIf(m_ValValore = "", "NULL", m_ValValore), ",", ".")
            result = result + UpdateStatus()
          End If
        Next
        'End If
      Loop
ErrUpdateSt:
      If (err.Number <> 0) Then ErrObj.WriteLog (" - Errore in import status- " & err.Description)
    End If
    .CloseFile
  End With
  Set ReaderObj = Nothing
  ImportStatus = IIf(result > 0, True, False)
  Exit Function
  
  
Err_StatusImport:
  If (CurErr > 0) Then
    ErrObj.WriteLog ("Errore in import status - " & err.Description)
    Resume Next
  Else
    ErrObj.WriteLog ("Errore in import status - " & err.Description)
    Set ReaderObj = Nothing
  End If
End Function

Private Function UpdateStatus() As Integer
  On Error GoTo Err_UpdateStatus
    With rsIns
        strSQL = " UPDATE " & .m_Tabella & " SET " & _
                 " DT_Status = " & .ParSQLStr(m_ValValore) & "" & _
                 " WHERE " & _
                 " DT_StationCode = " & .ParSQLStr(m_StationCode) & " AND " & _
                 " DT_MeasureCod = " & .ParSQLStr(m_ValParam) & " AND " & _
                 " DT_Date = " & .ParSQLStr(m_ValData) & " AND " & _
                 " DT_Hour = " & .ParSQLStr(m_ValOra)
                       
        .ExecuteSQL strSQL
    End With
      
'    If (AppServer2 <> "") Then
'      rsIns2.ExecuteSQL strSQL
'    End If
    
    UpdateStatus = 1
  
  Exit Function
Err_UpdateStatus:
  Resume Next
  'err.Raise err.Number, "UpdateStatus : ", err.Description & vbCrLf & strSQL
  'On Error GoTo 0
End Function

Public Function ImportMinMax(Filename As String, StationCode As String) As Integer
  Dim ReaderObj As Object
  Dim Intesta() As String
  Dim UniMis() As String
  Dim param() As String
  Dim i As Integer
  Dim CurErr As Integer
  Dim result As Long
  Dim MinStr As String
  Dim MaxStr As String
  Dim NumStr As String
  Dim anno As String
  Dim Table As String
  
  On Error GoTo Err_MinMaxImport
  ErrObj.Station = StationCode
  ErrObj.LookUpKey = "STATUS MINMAX"
  ErrObj.WriteLog "Import minmax"
  DoEvents
  m_StationCode = StationCode
  
  rsIns.m_Tabella = TABELLA
  Table = TABELLA
  
  Set ReaderObj = CreateObject("AttimoFwk.CFileReader")
  With ReaderObj
  
    If (.OpenFile(Filename)) Then
      On Error GoTo 0
      On Error GoTo ErrUpdateSt
      .SetDataMode .TabDelimited
      'Riga da saltare
      .ReadLine
      DoEvents
      'Riga intestazione contente i codici dei parametri
      .ReadLine
      ReDim Intesta(.ColumnCount)
      For i = 0 To .ColumnCount
        Intesta(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      .ReadLine
      ReDim UniMis(.ColumnCount)
      For i = 0 To .ColumnCount
        UniMis(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      .ReadLine
      ReDim param(.ColumnCount)
      For i = 0 To .ColumnCount
        param(i) = Trim$(.GetColumnValue(i))
        DoEvents
      Next
      
      'Riga da saltare
      '.ReadLine
         
      CurErr = 1
      Do While .ReadLine
        DoEvents
        m_ValOra = Replace(Trim$(.GetColumnValue(0)), ",", ".")
        'If (IsNumeric(m_ValOra)) Then
          'BFLAB Esporta le ore nel formato 1-24
          'BFDesk le gestisce nel formato 0-23
          'm_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")
          
        m_ValOra = Left$(Replace(m_ValOra, ":", "."), 5)
'        If (DataType <> Automatic) Then _
'          m_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")
          
        m_ValOra = Replace(Format$(CDbl(m_ValOra) - 1, "00.00"), ",", ".")

        m_ValData = Trim$(Replace(ReaderObj.GetColumnValue(1), "'", ""))
        m_ValData = Mid$(m_ValData, 7, 4) & Mid$(m_ValData, 4, 2) & Left$(m_ValData, 2)
        
        '21/02/2006
        anno = Left$(m_ValData, 4)
        If (rsIns.Table_Exist(Table & anno)) Then
          rsIns.m_Tabella = Table & anno
        Else
          rsIns.m_Tabella = Table
        End If
        
'        If (AppServer2 <> "") Then
'          If (rsIns2.Table_Exist(Table & Anno)) Then
'            rsIns2.m_Tabella = Table & Anno
'          Else
'            rsIns2.m_Tabella = Table
'          End If
'        End If
        
        '21/02/2006
      
        For i = 2 To .ColumnCount
          DoEvents
          If (Intesta(i) <> "") Then
            m_ValParam = Intesta(i)
            MinStr = Trim$(ReaderObj.GetColumnValue(i))
            MinStr = Replace(IIf(MinStr = "", "NULL", MinStr), ",", ".")
            
            MaxStr = Trim$(ReaderObj.GetColumnValue(i + 1))
            MaxStr = Replace(IIf(MaxStr = "", "NULL", MaxStr), ",", ".")
            
            NumStr = Trim$(ReaderObj.GetColumnValue(i + 2))
            NumStr = Replace(IIf(NumStr = "", "NULL", NumStr), ",", ".")
            
            result = result + UpdateMinMax(MinStr, MaxStr, NumStr)
            
            i = i + 2
          End If
        Next
        'End If
      Loop
ErrUpdateSt:
      If (err.Number <> 0) Then ErrObj.WriteLog (" - Errore in import status- " & err.Description)
    End If
    .CloseFile
  End With
  Set ReaderObj = Nothing
  ImportMinMax = IIf(result > 0, True, False)
  Exit Function
  
  
Err_MinMaxImport:
  If (CurErr > 0) Then
    ErrObj.WriteLog ("Errore in import status - " & err.Description)
    Resume Next
  Else
    ErrObj.WriteLog ("Errore in import status - " & err.Description)
    Set ReaderObj = Nothing
  End If
End Function

Private Function UpdateMinMax(MinStr As String, MaxStr As String, NumStr As String) As Integer
  On Error GoTo Err_UpdateStatus
  'If (m_ValValore <> "NULL") Then
    With rsIns
        strSQL = " UPDATE " & .m_Tabella & " SET " & _
                 " DT_Min = " & MinStr & "," & _
                 " DT_Max = " & MaxStr & "," & _
                 " DT_Nr = " & NumStr & _
                 " WHERE " & _
                 " DT_StationCode = " & .ParSQLStr(m_StationCode) & " AND " & _
                 " DT_MeasureCod = " & .ParSQLStr(m_ValParam) & " AND " & _
                 " DT_Date = " & .ParSQLStr(m_ValData) & " AND " & _
                 " DT_Hour = " & .ParSQLStr(m_ValOra)
                       
        .ExecuteSQL strSQL
    End With
    
'    If (AppServer2 <> "") Then
'      rsIns2.ExecuteSQL strSQL
'    End If
  
  'End If
  UpdateMinMax = 1
  
  Exit Function
Err_UpdateStatus:
  Resume Next
  'err.Raise err.Number, "UpdateMinMax : ", err.Description & vbCrLf & strSQL
  'On Error GoTo 0
End Function

'1.3.13


Private Sub TransformValue()
  Dim fattore As Double
  
  On Error GoTo Err_TransformValue
  With rsIns
    If (.Table_Exist("WDS_FACTOR")) Then
      If (.SelectionFast("SELECT FC_FACTOR,FC_UnitMeasure FROM WDS_FACTOR WHERE FC_STATIONCODE = " & _
          .ParSQLStr(m_StationCode) & " AND FC_MEASURECODE=" & .ParSQLStr(m_ValParam))) Then
          fattore = .GetValue("FC_FACTOR")
          m_ValValore = Replace(Trim(Str(Val(m_ValValore) * fattore)), ",", ".")
          If (.GetValue("FC_UnitMeasure") <> "") Then _
            m_ValUniMis = .GetValue("FC_UnitMeasure")
      End If
    End If
  End With
  
'  If (AppServer2 <> "") Then
'    With rsIns2
'      If (.Table_Exist("WDS_FACTOR")) Then
'        If (.SelectionFast("SELECT FC_FACTOR,FC_UnitMeasure FROM WDS_FACTOR WHERE FC_STATIONCODE = " & _
'            .ParSQLStr(m_StationCode) & " AND FC_MEASURECODE=" & .ParSQLStr(m_ValParam))) Then
'            fattore = .GetValue("FC_FACTOR")
'            m_ValValore = Replace(Trim(Str(Val(m_ValValore) * fattore)), ",", ".")
'            If (.GetValue("FC_UnitMeasure") <> "") Then _
'              m_ValUniMis = .GetValue("FC_UnitMeasure")
'        End If
'      End If
'    End With
'  End If
  
  Exit Sub
Err_TransformValue:
  Resume Next
End Sub


