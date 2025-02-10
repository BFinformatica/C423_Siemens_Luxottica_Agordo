Attribute VB_Name = "RecuperoDatidaPLC"
Option Explicit

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
'Global PLC_Connected As Boolean
Global PLC_Connected(10) As Boolean


Global Const timeout = 200

Global DBDati          As Integer
Global DWDati          As Integer
Global NBytes          As Integer
Global PLCData         As PLCDATA_DEF
Global CurrentPLC      As Integer


'******* matrici dati ******
Global AnalogReadings(256) As Integer
Global AnalogWritings(256) As Integer
Global DigitalReadings(256) As Integer
Global DigitalWritings(256) As Boolean

Global EffectiveAnalogWritings(64) As Integer
Global EffectiveDigitalWritings(64) As Integer


Global AnalogFromDCS(128) As Single
Global DigitalFromDCS(200) As Integer
Global AnalogToDCS(64) As Single
Global DigitalToDCS(256) As Integer

Global EffectiveAnalogToDCS(64) As Single
Global EffectiveDigitalToDCS(256) As Integer

Global LINEA As Integer
Global TG As Integer
Global IndirizzoIP(10) As String


Dim Anno As Integer
Dim Mese As Integer
Dim Giorno As Integer
Dim Ora As Integer
Dim Minuto As Integer
Dim Secondo As Integer
Dim valore(64) As Single

Dim ValoreDI(5) As Integer

Global Const PorteDigitali = 5





Sub PacchettoDati(numPLC, Pacchetti, DB)

    Dim lRetCode
    Dim Offset
    Dim iIdx As Integer
    
    On Error GoTo GestErrore

    DBDati = DB
    If DB = 20 Then
        Offset = 44 + (Pacchetti * 48)
    Else
        Offset = Pacchetti * 48
    End If
    
    'Letture Data e ora
    DWDati = 0 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Anno)
    
    DWDati = 1 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Mese)
    
    DWDati = 2 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Giorno)
    
    DWDati = 3 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Ora)
    
    DWDati = 4 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Minuto)
    
    DWDati = 5 + Offset
    NBytes = 1   'N. bytes letti / 8 ingressi da 4 byte
    lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, Secondo)
    
    If Not (Mese = 0 And Giorno = 0 And Anno = 0) Then
    
        'TimeStamp = CDate(Format(Mese, "00") + "/" + Format(Giorno, "00") + "/20" + Format(Anno, "00") + " " + Format(Ora, "00") + ":" + Format(Minuto, "00") + ":" + Format(Secondo, "00"))
        
        'Letture registri storicizzati
        DWDati = 6 + Offset
        NBytes = (4 * 9)     'N. bytes letti / 8 ingressi da 4 byte
        
        lRetCode = PLCDWordRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, valore(0))
        
        
        'Lettura Word rappresentativa dei bit acquisiti
        For iIdx = 0 To PorteDigitali
            DWDati = 42 + iIdx + Offset
            NBytes = 1      'N. bytes letti / 3 ingressi da 2 byte
            
            lRetCode = PLCByteRead(rgPLCDef(numPLC).hPLCHandle, DBDati, DWDati, NBytes, ValoreDI(iIdx))
        Next
        
        'Call ElaboraDatiRecuperodaPLC
    
    End If
    
    Exit Sub
    
GestErrore:
    
    Call WindasLog("PacchettoDati errore: " + Error(Err), 1, OPC)
    Resume Next
End Sub


Sub InizializzaProtocollo(nn As Integer)
    
    'Alby Giugno 2016
    
    On Error GoTo GestErr
    
    Dim PingObj As Object
    
    '****** preimposta stato ******
    PLC_Connected(nn) = True
    rgPLCDef(nn).bInUso = True

    
    If ReadFileIni Then
        
        IndirizzoIP(nn) = rgPLCDef(nn).iIPAddress(1) & "." & rgPLCDef(nn).iIPAddress(2) & "." & rgPLCDef(nn).iIPAddress(3) & "." & rgPLCDef(nn).iIPAddress(4)
        
        Set PingObj = CreateObject("AttimoFwk.CPing")
            
        If PingObj.Ping(IndirizzoIP(nn)) Then
            If Not ConnectPLC(nn) Then
                warn_msg = "Setup Comunication with PLC: Connection refused"
                PLC_Connected(nn) = False
            End If
        Else
            warn_msg = "Setup Comunication with PLC: ping failed!"
            PLC_Connected(nn) = False
        End If
        
    
    Else
        warn_msg = "Setup Comunication with PLC: File .INI error"
        PLC_Connected(nn) = False
    End If

    Exit Sub
    
GestErr:
    Call WindasLog("InizializzaProtocollo " + Error(Err), 1, OPC)

End Sub






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

Function AssegnaPorta(Porta)

    Select Case Porta
            
        Case 0, 1
            AssegnaPorta = 0
        
        Case 2, 3
            AssegnaPorta = 1
    
        Case 4, 5
            AssegnaPorta = 2
    
    End Select

End Function

Function AssegnaBit(Porta, NrBit)

    Select Case Porta
            
        Case 0, 2, 4
            AssegnaBit = NrBit + 8
    
        Case 1, 3, 5
            AssegnaBit = NrBit
    
    
    End Select

End Function
