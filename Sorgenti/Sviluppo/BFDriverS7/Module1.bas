Attribute VB_Name = "Module1"
Global Const ERR_LOG = 0
Global Const MSG_LOG = 1

Global frequenza As Long

Global IdStrumento As String
Global InstrumentType As String


Global MeasureStatus As String
Global no_com_ctr As Integer
Global time_out_ctr As Integer
Global str_version As String
Global warn_msg As String
Global CRLF As String * 2
Global OutputCommand As String
Global PLC_Connected As Boolean

Global Const TIMEOUT = 200

Global DBDati          As Integer
Global DWDati          As Integer
Global NBytes          As Integer
Global PLCData         As PLCDATA_DEF
Global CurrentPLC      As Integer

'******* matrici dati ******
Global AnalogReadings(256) As Single
Global iAnalogWritings(256) As Integer
Global fAnalogWritings(256) As Single
Global DigitalReadings(256) As Single
Global DigitalWritings(256) As Integer

Global iEffectiveAnalogWritings(256) As Integer
Global fEffectiveAnalogWritings(256) As Integer
Global EffectiveDigitalWritings(256) As Integer

Global iMaxAI As Integer
Global iMaxDI As Integer
Global iMaxAO As Integer
Global iMaxDO As Integer

Global AnalogFromDCS(256) As Single
Global DigitalFromDCS(256) As Integer
Global AnalogToDCS(256) As Single
Global DigitalToDCS(256) As Integer

Global EffectiveAnalogToDCS(256) As Single
Global EffectiveDigitalToDCS(256) As Integer

Global linea As Integer
Global LineaBFlab As String
Global TG As Integer

Global NoDigOut As Boolean

'*** michele - ping
Global PingObj As Object
Global IP_Master As String

Global BFComunicator As Object

'*** utilizzati per la parametrizzazione delle letture analogiche
Global MappaturaAI(127, 7) As String
Global Parametri As Variant
Global mRecordsetCols As Long
Global mRecordCount As Long

'*** utilizzati per la parametrizzazione delle letture analogiche
Global MappaturaDI(127, 4) As String
Global ParametriDI As Variant
Global mRecordsetColsDI As Long
Global mRecordCountDI As Long

'*** utilizzati per la parametrizzazione delle uscite analogiche
Global MappaturaAO(127, 7) As String
Global ParametriAO As Variant
Global mRecordsetColsAO As Long
Global mRecordCountAO As Long

'*** utilizzati per la parametrizzazione delle uscite digitali
Global MappaturaDO(127, 10) As String
Global ParametriDO As Variant
Global mRecordsetColsDO As Long
Global mRecordCountDO As Long

Global AppTry As CTryArea
Global close_cmd As Boolean
Global hide_cmd As Boolean

Global BF_Driver As Object

Global Const TIPO_INTEGER = 0
Global Const TIPO_FLOAT = 1
Global Const ANALOGICI = 0
Global Const DIGITALI = 1
    
Global BFComunicatorDisabled As Boolean
Global bComunicationAlarmEnabled As Boolean
Global iComunicationAlarmIndex As Integer
Global bWatchDogEnabled As Boolean
Global sWD_TAG As String


Sub Main()

    'Alby Marzo 2018
    If App.PrevInstance Then End

    Set BF_Driver = New BFserver
    BF_Driver.StartMe
  
End Sub


Public Sub Ritardo(ByVal fSecondi As Single)

    Dim dAttesa As Double
    
    Rem ***** Ciclo di Ritardo *****
    dAttesa = Timer
    Do
        DoEvents
    Loop Until Mezzanotte(dAttesa) > fSecondi

End Sub
Public Function Mezzanotte(ByVal dAttesa As Double) As Double
    
    Dim dLetto As Double
    
    dLetto = Timer
    Mezzanotte = dLetto - dAttesa
    
    If CLng(Mezzanotte) < 0 Then
        Mezzanotte = 86400 - dAttesa + dLetto
    End If

End Function

Public Sub SalvaLog(ByVal tipo_log As Integer, ByVal msg As String)

    Dim ll As Integer
    Dim log_path As String
    Dim log_file As String
    
    On Error Resume Next
    
    '***** crea cartella se non esiste ******
    log_path = App.Path & "\DriverLog"
    If (Dir(log_path, vbDirectory) = "") Then MkDir (log_path)
    
    '********* seleziona file log ***********
    Select Case tipo_log
        Case ERR_LOG
            log_file = log_path & "\BFDriver_error.log"
        Case MSG_LOG
            log_file = log_path & "\BFDriver_msg.log"
    End Select
    
    ll = FreeFile
    Open log_file For Append As #ll
    Print #ll, Format(Now, "dd/mm/yyyy hh.nn.ss") & Chr(9) & msg
    Close (ll)

End Sub


