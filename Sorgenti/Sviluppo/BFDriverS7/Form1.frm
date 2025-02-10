VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BF Driver S7 per PLC SIEMENS"
   ClientHeight    =   9450
   ClientLeft      =   5130
   ClientTop       =   4995
   ClientWidth     =   15510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmd_mappa 
      Caption         =   "Mappatura I/O"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   8940
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ingressi Analogici"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txt_AI(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txt_AI(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txt_AI(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ingressi Digitali"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_DI(0)"
      Tab(1).Control(1)=   "txt_DI(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Uscite analogiche"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_AO(1)"
      Tab(2).Control(1)=   "txt_AO(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Uscite digitali"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt_DO(1)"
      Tab(3).Control(1)=   "txt_DO(0)"
      Tab(3).ControlCount=   2
      Begin VB.TextBox txt_AI 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7860
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   480
         Width           =   4860
      End
      Begin VB.TextBox txt_AO 
         Height          =   7860
         Index           =   0
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_DO 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7860
         Index           =   0
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_AI 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7860
         Index           =   1
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   480
         Width           =   4860
      End
      Begin VB.TextBox txt_DI 
         Height          =   7860
         Index           =   1
         Left            =   -67320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_DI 
         Height          =   7860
         Index           =   0
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_AO 
         Height          =   7860
         Index           =   1
         Left            =   -67320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_DO 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7860
         Index           =   1
         Left            =   -67320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   480
         Width           =   7380
      End
      Begin VB.TextBox txt_AI 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7860
         Index           =   2
         Left            =   10200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   480
         Width           =   4860
      End
   End
   Begin VB.Label lbl_warnings 
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1575
      TabIndex        =   1
      Top             =   8940
      Width           =   8595
   End
   Begin VB.Label lbl_versione 
      Alignment       =   1  'Right Justify
      Caption         =   "Versione"
      Height          =   255
      Left            =   10440
      TabIndex        =   0
      Top             =   9060
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub InizializzaProtocollo()
    
    Dim nn As Integer
    
    
    On Error GoTo GestErr
    
    '****** preimposta stato ******
    PLC_Connected = True
    
    'da levare
    'Exit Sub
    
    If ReadFileIni Then
      CurrentPLC = ID_MASTER
      
      '**********************************************************************
      '*                       Check della connessione                      *
      '**********************************************************************
      If PingObj.Ping(IP_Master) Then
        warn_msg = ""
        PLC_Connected = True
        
        If Not ConnectPLC(CurrentPLC) Then
          warn_msg = "InizializzaProtocollo: connessione NON riuscita con PLC MASTER!"
          PLC_Connected = False
        End If
        
      Else
        warn_msg = "PLC SCOLLEGATO!"
        PLC_Connected = False
        Exit Sub
      End If

    Else
    
      warn_msg = "InizializzaProtocollo: errore in lettura file .INI!"
      
    End If
    
    '*** warnings ****
    lbl_warnings.Caption = warn_msg

    Exit Sub
    
GestErr:
    warn_msg = "InizializzaProtocollo: " & Err.Description
    Debug.Print Err.Description
    Call SalvaLog(ERR_LOG, "InizializzaProtocollo: " & Err.Description)
    Resume Next

End Sub

Private Sub LeggiRegistriDigitali()

    Dim di_ndx As Integer
    Dim iIdx As Integer
    Dim ndx As Integer
    Dim testo As String
    
    Dim bit As Integer
    Dim lRetCode As Long
    Dim DigValore(255) As Byte
    Dim bError As Boolean
    
    On Error GoTo GestErrore
    
    '*** reset display
    testo = ""
    txt_DI(0).Text = ""
    txt_DI(1).Text = ""
    di_ndx = -1
    
    If Not PLC_Connected Then GoTo RiportoComAlrm
    
    '**********************************************************************
    '*                        INGRESSI DIGITALI                           *
    '**********************************************************************
    
    warn_msg = ""
    
    '*** reset indice dI
    di_ndx = 0
    
    '*** reset display
    testo = ""
    txt_DI(0).Text = ""
    txt_DI(1).Text = ""
    
    '**** scorre tutte le righe di lettura DI impostate in configurazione
    For iIdx = 0 To mRecordCountDI - 1
    
        DBDati = Val(MappaturaDI(iIdx, 1))      'N. DB
        DWDati = Val(MappaturaDI(iIdx, 2))      'offset
        NBytes = Val(MappaturaDI(iIdx, 3))      'N. bytes letti
        
        'da levare commento
        lRetCode = PLCByteRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, DigValore(0))
        
        'da levare
        'lRetCode = 1

        For ndx = 0 To NBytes - 1
            
            For bit = 0 To 7
            
                If (lRetCode <> 0) Then
                    If (DigValore(ndx) And 2 ^ bit) = 2 ^ bit Then
                      DigitalReadings(di_ndx) = 1
                    Else
                      DigitalReadings(di_ndx) = 0
                    End If
                Else
                    bError = True
                End If
              
                '*****************************************************************************
                '*   gestione display: 0..38 sul primo text box, >=48 sul secondo textbox    *
                '*****************************************************************************
                If di_ndx = 39 Then
                    txt_DI(0).Text = testo
                    testo = ""
                ElseIf di_ndx = 78 Then
                    txt_DI(1).Text = testo
                    testo = ""
                End If
                
                testo = testo & Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & di_ndx
                testo = testo & "   " & MappaturaDI(iIdx, 0)
                testo = testo & "   " & "DB" & MappaturaDI(iIdx, 1)
                testo = testo & "." & (DWDati + ndx) & "." & bit
                testo = testo & " = " & DigitalReadings(di_ndx) & CRLF
                '*****************************************************************************
                
                '**** punta all'ingresso successivo
                di_ndx = di_ndx + 1
                
            Next bit
        
        Next ndx
        
    Next iIdx
    
    If bError Then
      If warn_msg = "" Then
        warn_msg = "Errore di lettura DI!"
      Else
        warn_msg = warn_msg & " \ Errore di lettura DI!"
      End If
    End If

RiportoComAlrm:

    '**** riporta l'eventuale allarme di comunicazione
    If bComunicationAlarmEnabled Then
        testo = testo & Format(Now, "dd/mm/yyyy hh:nn:ss") & "  PING col PLC - DI" & iComunicationAlarmIndex
        testo = testo & " = " & DigitalReadings(iComunicationAlarmIndex)
    End If

    '**** visualizzazione su text box
    Select Case di_ndx
      Case 0 To 38
        txt_DI(0).Text = testo
      Case 33 To 77
        txt_DI(1).Text = testo
      Case Else
    End Select
    
    iMaxDI = di_ndx
    If iMaxDI > 255 Then iMaxDI = 255
    Call AggiornaBFComunicator(DIGITALI)

    Exit Sub

GestErrore:
    warn_msg = "LeggiRegistriDigitali: " & Error(Err)
    Debug.Print warn_msg
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Sub

Private Sub ScriviRegistriAnalogici()

    Dim ao_ndx As Integer
    Dim iIdx As Integer
    Dim ndx As Integer
    Dim tipo_var As Integer
    Dim bytes_per_dato As Integer
    Dim testo As String
    
    Dim start_ao As Integer
    Dim end_ao As Integer
    Dim range_ao As String
    Dim dati() As String
    Dim iUscitaAO(256) As Integer
    Dim fUscitaAO(256) As Single
    Dim bError As Boolean
    
    Dim lRetCode As Long
    Dim lRetCode1 As Long
    
    On Error GoTo GestErrore
    
    '**********************************************************************
    '*                        USCITE ANALOGICHE                           *
    '**********************************************************************
    'ParametriAO(0, riga) = Descrizione
    'ParametriAO(1, riga) = Numero DB
    'ParametriAO(2, riga) = Indirizzo base (OFFSET)
    'ParametriAO(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'ParametriAO(4, riga) = N. uscite da scrivere
    'ParametriAO(5, riga) = N. bytes da scrivere (valore calcolato)
    'ParametriAO(6, riga) = Range uscite (valore calcolato)

    '*** reset indice AI
    ao_ndx = 0
    testo = ""
    
    '**** se il PLC non risponde al PING...
    If Not PLC_Connected Then Exit Sub
        
    '**** legge da BFComunicator le tag del tipo "[linea] AOx"
    Call LeggiBFComunicator(ANALOGICI)

    For iIdx = 0 To mRecordCountAO - 1
    
        DBDati = Val(MappaturaAO(iIdx, 1))      'N. DB
        DWDati = Val(MappaturaAO(iIdx, 2))      'offset
        tipo_var = Val(MappaturaAO(iIdx, 3))    '0 = integer   1 = float
        NBytes = Val(MappaturaAO(iIdx, 5))      'N. bytes da scrivere
        bytes_per_dato = IIf(tipo_var = TIPO_INTEGER, 2, 4)
        
        'michele sett 2014: determina il range di indici delle uscite
        range_ao = MappaturaAO(iIdx, 6)
        dati = Split(range_ao, "-")
        If UBound(dati) = 1 Then
          start_ao = Val(Trim(dati(0)))
          end_ao = Val(Trim(dati(1)))
          
          If tipo_var = TIPO_INTEGER Then
          
            'trasferisce al vettore temporaneo (il cui indice è a base 0) le uscite lette dal BFComunicator
            For ndx = start_ao To end_ao
              iUscitaAO(ndx - start_ao) = iAnalogWritings(ndx)
            Next ndx
            
            'da levare commento
            lRetCode = PLCWordWrite(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, iUscitaAO(0))
          
            '**** rilegge i registro per conferma avvenuta scrittura
            'da levare commento
            lRetCode1 = PLCWordRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, iUscitaAO(0))
            
            'trasferisce dal vettore temporaneo (il cui indice è a base 0) le uscite lette dal PLC
            For ndx = start_ao To end_ao
              iEffectiveAnalogWritings(ndx) = iUscitaAO(ndx - start_ao)
            Next ndx
          
          Else
          
            'trasferisce al vettore temporaneo (il cui indice è a base 0) le uscite lette dal BFComunicator
            For ndx = start_ao To end_ao
              fUscitaAO(ndx - start_ao) = fAnalogWritings(ndx)
            Next ndx
            
            'da levare commento
            lRetCode = PLCDWordWrite(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, fUscitaAO(0))
          
            '**** rilegge i registro per conferma avvenuta scrittura
            'da levare commento
            lRetCode1 = PLCDWordRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, fUscitaAO(0))
            
            'trasferisce dal vettore temporaneo (il cui indice è a base 0) le uscite lette dal PLC
            For ndx = start_ao To end_ao
              fEffectiveAnalogWritings(ndx) = fUscitaAO(ndx - start_ao)
            Next ndx
          
          End If
          
          '**** setta errore di scrittura
          If lRetCode = 0 Then bError = True
            
        End If

        '*****************************************************************************
        '*   gestione display: 0..31 sul primo text box, >=32 sul secondo textbox    *
        '*****************************************************************************
        DWDati = Val(MappaturaAO(iIdx, 2))      'offset
        For ndx = start_ao To end_ao

            If ao_ndx = 32 Then
                txt_AO(0).Text = testo
                testo = ""
            End If

            testo = testo & Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & ao_ndx
            testo = testo & "   " & MappaturaAO(iIdx, 0)
            testo = testo & "   " & "DB" & MappaturaAO(iIdx, 1)
            testo = testo & "." & (DWDati + (ndx - start_ao) * bytes_per_dato)
            
            If tipo_var = TIPO_INTEGER Then
                testo = testo & " = " & iAnalogWritings(ao_ndx) & " (check: " & iEffectiveAnalogWritings(ndx) & ")" & CRLF
            Else
                testo = testo & " = " & fAnalogWritings(ao_ndx) & " (check: " & fEffectiveAnalogWritings(ndx) & ")" & CRLF
            End If

            '**** punta all'ingresso successivo
            ao_ndx = ao_ndx + 1
            
        Next ndx

    Next iIdx
    
    '**** visualizzazione su text box
    If ao_ndx < 32 Then
        txt_AO(0).Text = testo
        txt_AO(1).Text = ""
    Else
        txt_AO(1).Text = testo
    End If
    
    iMaxAO = ao_ndx

    If bError Then
      If warn_msg = "" Then
        warn_msg = "Errore di scrittura AO!"
      Else
        warn_msg = warn_msg & " \ Errore di scrittura AO!"
      End If
    End If
    
    Exit Sub

GestErrore:
    warn_msg = "ScriviRegistriAnalogici: " & Error(Err)
    Debug.Print warn_msg
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Sub

Private Sub ScriviRegistriDigitali()

    Dim do_ndx As Integer

    Dim iIdx As Integer
    Dim ndx As Integer
    Dim testo As String
    Dim bit As Integer
    Dim lRetCode As Long
    Dim DigValore(255) As Byte
    Dim bError As Boolean
    
    On Error GoTo GestErrore
    
    '**********************************************************************
    '*                          USCITE DIGITALI                           *
    '**********************************************************************
    '*** reset indice DO
    do_ndx = 0

    '*** reset display
    testo = ""
    
    '**** se il PLC non risponde al PING...
    If Not PLC_Connected Then Exit Sub

    '**** legge da BFComunicator le tag del tipo "[linea] DOx"
    Call LeggiBFComunicator(DIGITALI)
    
    For iIdx = 0 To mRecordCountDO - 1
    
        DBDati = Val(MappaturaDO(iIdx, 1))      'N. DB
        DWDati = Val(MappaturaDO(iIdx, 2))      'offset
        NBytes = Val(MappaturaDO(iIdx, 3))      'N. bytes letti
        
        For ndx = 0 To NBytes - 1
        
            '***** resetta il byte da spedire
            DigValore(0) = 0
            
            For bit = 0 To 7
            
                '**** compone il byte da spedire
                If DigitalWritings(do_ndx) = 1 Then
                    DigValore(0) = DigValore(0) Or 2 ^ bit
                End If
                
                '**** punta all'uscita successiva
                do_ndx = do_ndx + 1
                
            Next bit
            
            '***** invia il byte di 8 uscite
            lRetCode = PLCByteWrite(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, 1, DigValore(0))
            
            '***** gestione errori di scrittura
            If lRetCode = 0 Then bError = True
            
            '***** rilegge il byte per conferma avvenuto invio
            lRetCode = PLCByteRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, DigValore(0))
            
            '***** riporta indietro il contatore all'inizio del byte
            do_ndx = do_ndx - 8
            
            '*****************************************************************************
            '*   gestione display: 0..31 sul primo text box, >=32 sul secondo textbox    *
            '*****************************************************************************
            For bit = 0 To 7
            
                If (lRetCode <> 0) Then
                    If (DigValore(0) And 2 ^ bit) = 2 ^ bit Then
                      EffectiveDigitalWritings(do_ndx) = 1
                    Else
                      EffectiveDigitalWritings(do_ndx) = 0
                    End If
                End If
                
                If do_ndx = 40 Then
                    txt_DO(0).Text = testo
                    testo = ""
                End If
                
                testo = testo & Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & do_ndx
                testo = testo & "   " & MappaturaDO(iIdx, 0)
                testo = testo & "   " & "DB" & MappaturaDO(iIdx, 1)
                testo = testo & "." & (DWDati) & "." & bit
                testo = testo & " = " & DigitalWritings(do_ndx) & " (check: " & EffectiveDigitalWritings(do_ndx) & ")" & CRLF
                '*****************************************************************************

                '**** punta all'uscita successiva
                do_ndx = do_ndx + 1

            Next bit
            
            DWDati = DWDati + 1
                
        Next ndx
        
    Next iIdx

    '**** visualizzazione su text box
    If do_ndx < 40 Then
        txt_DO(0).Text = testo
        txt_DO(1).Text = ""
    Else
        txt_DO(1).Text = testo
    End If
    
    iMaxDO = do_ndx
    
    If bError Then
      If warn_msg = "" Then
        warn_msg = "Errore di scrittura DO!"
      Else
        warn_msg = warn_msg & " \ Errore di scrittura DO!"
      End If
    End If
    
    Exit Sub

GestErrore:
    warn_msg = "ScriviRegistriDigitali: " & Error(Err)
    Call SalvaLog(ERR_LOG, warn_msg)
    Debug.Print warn_msg
    Resume Next

End Sub


Public Sub Terminate()
  End
End Sub

Private Sub VerificaConnessione()

    On Error Resume Next
    
    If PingObj.Ping(IP_Master) Then
      warn_msg = ""
      PLC_Connected = True
    Else
      warn_msg = "PLC SCOLLEGATO!"
      PLC_Connected = False
    End If

End Sub


Private Sub cmd_mappa_Click()

    Timer1.Enabled = False
    Form2.Show
    
End Sub


Private Sub Form_Load()
    
    lbl_versione.Caption = str_version
    
    CRLF = Chr(13) & Chr(10)
    
    Set AppTry = New CTryArea
    
End Sub

Private Sub Form_Activate()

    On Error Resume Next
    
    Me.Hide
    AppTry.IconAdd Me, App.ProductName

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frm_closing.Show 1
    If hide_cmd Then
      Cancel = 1
      Me.Hide
      Exit Sub
    Else
      If MsgBox("Confermi la chiusura di BFDriver_S7?", vbOKCancel Or vbQuestion, "") <> vbOK Then
        Cancel = 1
        Exit Sub
      End If
    End If
    
    AppTry.IconDelete
    Set AppTry = Nothing
    Set BF_Driver = Nothing
    Unload Me
    End

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'this procedure receives the callbacks from the System Tray icon.
    Dim result As Long
    Dim msg As Long
    Const WM_LBUTTONDOWN = &H201     'Button down
    Const WM_LBUTTONUP = &H202       'Button up
    Const WM_LBUTTONDBLCLK = &H203   'Double-click
    
    On Error Resume Next
    
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    msg = X / Screen.TwipsPerPixelX
    
    Select Case msg
        Case WM_LBUTTONUP, WM_LBUTTONDBLCLK        '514 restore form window
          AppTry.RestoreWindow
          Me.WindowState = vbNormal
          Me.Show
          'Me.ZOrder -1
        'Case WM_RBUTTONUP        '517 display popup menu
        ' Result = SetForegroundWindow(Me.hwnd)
        ' Me.PopupMenu Me.mPopupSys
    End Select
    
End Sub

Private Sub Timer1_Timer()

    Static iStato As Integer
    Static iCount As Integer
    
    Const INIT = 0
    Const COLLEGA = 1
    Const ACQUISISCE = 2
    
    On Error Resume Next
    
    Select Case iStato
      Case INIT
        Me.Caption = "In partenza..."
        Me.Refresh
        If iCount < 3 Then iCount = iCount + 1
        If iCount = 3 Then iStato = COLLEGA
        
      Case COLLEGA
        Me.Caption = "In attesa di collegamento col PLC..."
        Call InizializzaProtocollo
        Me.Caption = "In attesa di collegamento col PLC... (" & IP_Master & ")"
        Me.Refresh
        If Not PLC_Connected Then
          Call Ritardo(1)
        Else
          Me.Caption = "BFDriver per PLC Siemens (" & IP_Master & ")"
          Me.Refresh
          Call SalvaLog(MSG_LOG, "BFDriver_S7 riavviato.")
          Call SalvaLog(MSG_LOG, "Connessione stabilita col PLC (" & IP_Master & ")")
          iStato = ACQUISISCE
        End If
        
        '***** per gestire il salvataggio del watchdog anche se non è ancora collegato...
        Call GestioneWatchDog

      Case ACQUISISCE
        Call AcquisiscePLC
        
    End Select
    
End Sub

Sub AcquisiscePLC()

    On Error Resume Next

    '**********************************************************************
    '*                   visualizzazione messaggistica                    *
    '**********************************************************************
    warn_msg = ""
  
    '**********************************************************************
    '*                       Check della connessione                      *
    '**********************************************************************
    Call VerificaConnessione
    
    '**********************************************************************
    '*                       INGRESSI ANALOGICI                           *
    '**********************************************************************
    Call LeggiRegistriAnalogici
    
    '**********************************************************************
    '*                        INGRESSI DIGITALI                           *
    '**********************************************************************
    Call LeggiRegistriDigitali

    '**********************************************************************
    '*                        USCITE ANALOGICHE                           *
    '**********************************************************************
    Call ScriviRegistriAnalogici

    '**********************************************************************
    '*                          USCITE DIGITALI                           *
    '**********************************************************************
    Call ScriviRegistriDigitali
    
    '**********************************************************************
    '*          Visualizza gli eventuali allarmi di comunicazione         *
    '**********************************************************************
    Call GestioneAllarmiComunicazione
    
    '**********************************************************************
    '*              Se impostato, salva una tag di watchdog               *
    '**********************************************************************
    Call GestioneWatchDog
    
    '**********************************************************************
    '*                   visualizzazione messaggistica                    *
    '**********************************************************************
    lbl_warnings.Caption = warn_msg

End Sub

Private Sub LeggiRegistriAnalogici()

    Dim ai_ndx As Integer

    Dim iIdx As Integer
    Dim ndx As Integer
    Dim tipo_var As Integer
    Dim bytes_per_dato As Integer
    Dim max_ai As Integer
    Dim testo As String
    
    Dim lRetCode As Long
    Dim IntValore(255) As Integer
    Dim FloatValore(255) As Single
    
    On Error GoTo GestErrore
    
    '*** reset display
    testo = ""
    
    '**** se il PLC non risponde al PING...
    If Not PLC_Connected Then Exit Sub
    
    '**********************************************************************
    '*                       REGISTRI ANALOGICI                           *
    '**********************************************************************
    'Parametri(0, riga) = Descrizione
    'Parametri(1, riga) = Numero DB
    'Parametri(2, riga) = Indirizzo base (OFFSET)
    'Parametri(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'Parametri(4, riga) = N. ingressi da leggere
    'Parametri(5, riga) = N. bytes da leggere (valore calcolato)
    'Parametri(6, riga) = Range Ingressi (valore calcolato)

    '*** reset indice AI
    ai_ndx = 0
    testo = ""
    
    '**** scorre tutte le righe di lettura AI impostate in configurazione
    For iIdx = 0 To mRecordCount - 1
    
        DBDati = Val(MappaturaAI(iIdx, 1))      'N. DB
        DWDati = Val(MappaturaAI(iIdx, 2))      'Indirizzo base (OFFSET)
        tipo_var = Val(MappaturaAI(iIdx, 3))    'Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
        max_ai = Val(MappaturaAI(iIdx, 4)) - 1  'N. ingressi da leggere
        NBytes = Val(MappaturaAI(iIdx, 5))      'N. bytes letti
        bytes_per_dato = IIf(tipo_var = TIPO_INTEGER, 2, 4)

        If tipo_var = TIPO_INTEGER Then
            'da levare commento
            lRetCode = PLCWordRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, IntValore(0))
        Else
            'da levare commento
            lRetCode = PLCDWordRead(rgPLCDef(CurrentPLC).hPLCHandle, DBDati, DWDati, NBytes, FloatValore(0))
        End If
        
        'da levare
        'lRetCode = 1

        For ndx = 0 To max_ai

            '**** Aggiorna la lettura AnalogReadings se la lettura dal PLC è andata a buon fine
            If (lRetCode <> 0) Then
                If tipo_var = TIPO_INTEGER Then
                    AnalogReadings(ai_ndx) = IntValore(ndx)
                ElseIf tipo_var = TIPO_FLOAT Then
                    AnalogReadings(ai_ndx) = FloatValore(ndx)
                End If
            Else
                warn_msg = "Errore di lettura AI!"
            End If

            '*****************************************************************************
            '*   gestione display: 0..39 sul primo text box, >=48 sul secondo textbox    *
            '*****************************************************************************
            If ai_ndx = 40 Then
                txt_AI(0).Text = testo
                testo = ""
            End If
            If ai_ndx = 80 Then
                txt_AI(1).Text = testo
                testo = ""
            End If

            testo = testo & Format(Now, "dd/mm/yyyy hh:nn:ss") & "  " & ai_ndx
            testo = testo & "   " & MappaturaAI(iIdx, 0)
            testo = testo & "   " & "DB" & MappaturaAI(iIdx, 1)
            testo = testo & "." & (DWDati + ndx * bytes_per_dato)
            testo = testo & " = " & AnalogReadings(ai_ndx) & CRLF
            '*****************************************************************************

            '**** punta all'ingresso successivo
            ai_ndx = ai_ndx + 1
            
        Next ndx

    Next iIdx
    
    '**** visualizzazione su text box
    If ai_ndx < 40 Then
        txt_AI(0).Text = testo
        txt_AI(1).Text = ""
        txt_AI(2).Text = ""
    ElseIf (ai_ndx >= 40) And (ai_ndx < 80) Then
        txt_AI(1).Text = testo
        txt_AI(2).Text = ""
    Else
        txt_AI(2).Text = testo
    End If

    iMaxAI = ai_ndx
    If iMaxAI > 255 Then iMaxAI = 255
    Call AggiornaBFComunicator(ANALOGICI)
    
    Exit Sub
    
GestErrore:
    warn_msg = "LeggiRegistriAnalogici: " & Error(Err)
    Debug.Print warn_msg
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Sub

Private Sub AggiornaBFComunicator(ByVal tipo_misura As Integer)

    Dim ai As Integer
    Dim di As Integer
    Dim ndx As Integer
    Dim Valore As Single
    Dim NomeTag As String
    Dim S As String
    Dim IndiceQAL3 As Integer
    
    Const ANALOGICI = 0
    Const DIGITALI = 1
    
    On Error GoTo GestErr
    
    '**** non aggiorna BFComunicator, se impostato
    If BFComunicatorDisabled Then Exit Sub
    
    If Not BFComunicator Is Nothing Then
      If tipo_misura = ANALOGICI Then

        For ndx = 0 To iMaxAI
            Call DatoComunicatorMemorizza(Trim("AI" & ndx), AnalogReadings(ndx))
        Next ndx
        
        Form2.GridEX_AI.MoveFirst
        
        For ndx = 1 To Form2.GridEX_AI.RowCount
        
            If InStr(UCase(Form2.GridEX_AI.Value(1)), "QAL3") > 0 Then
                NomeTag = UCase(Form2.GridEX_AI.Value(1))
                S = Trim(Left(Form2.GridEX_AI.Value(7), InStr(Form2.GridEX_AI.Value(7), "-") - 1))
                IndiceQAL3 = Val(S)
                Call DatoComunicatorMemorizza(Trim(NomeTag), AnalogReadings(IndiceQAL3))
            End If
            
            Form2.GridEX_AI.MoveNext
            
        Next ndx
        
      ElseIf tipo_misura = DIGITALI Then
        
        For ndx = 0 To iMaxDI
            Call DatoComunicatorMemorizza(Trim("DI" & ndx), DigitalReadings(ndx))
        Next ndx
        
        '**** salva eventuale allarme di comunicazione
        If bComunicationAlarmEnabled Then
            ndx = iComunicationAlarmIndex
            Call DatoComunicatorMemorizza(Trim("DI" & ndx), DigitalReadings(ndx))
        End If
        
      End If
    End If
    
    Exit Sub
    
GestErr:
    warn_msg = "AggiornaBFComunicator: " & Err.Description
    Err.Clear
    Resume Next

End Sub
Sub DatoComunicatorMemorizza(Item As String, Value)

    On Error GoTo GestErrore

    If Val(LineaBFlab) > 0 Then
      BFComunicator.AddItem LineaBFlab + " " + Item
      BFComunicator.CurrentItem = LineaBFlab + " " + Item
    Else
      BFComunicator.AddItem Item
      BFComunicator.CurrentItem = Item
    End If
    BFComunicator.ItemValue = Value
        
    Exit Sub

GestErrore:
    warn_msg = "DatoComunicatorMemorizza: " & Err.Description
    Err.Clear
    Resume Next

End Sub

Sub GestioneAllarmiComunicazione()

    Dim testo_allarmi As String
    
    On Error Resume Next
    
    '***** gestione del salvataggio degli allarmi di comunicazione
    testo_allarmi = ""
    If bComunicationAlarmEnabled Then
        DigitalReadings(iComunicationAlarmIndex) = IIf(PLC_Connected, 0, 1)
    End If

End Sub


Sub GestioneWatchDog()

    On Error GoTo GestErr
    
    If Not bWatchDogEnabled Then Exit Sub
    If sWD_TAG = "" Then Exit Sub
    Call DatoComunicatorMemorizza(sWD_TAG, Format(Timer, "0"))
    Exit Sub
    
GestErr:
    warn_msg = "GestioneWatchDog: " & Err.Description
    Err.Clear
    Resume Next
    
End Sub

Private Sub LeggiBFComunicator(ByVal tipo_misura As Integer)

    Dim ai As Integer
    Dim di As Integer
    Dim ndx As Integer
    Dim Valore As Single
    Dim testo As String
    
    Const ANALOGICI = 0
    Const DIGITALI = 1
    
    On Error GoTo GestErr
    
    If Not BFComunicator Is Nothing Then
      If tipo_misura = ANALOGICI Then
      
        For ndx = 0 To iMaxAO
          If Val(LineaBFlab) > 0 Then
            BFComunicator.CurrentItem = LineaBFlab & " " & "AO" & ndx
          Else
            BFComunicator.CurrentItem = "AO" & ndx
          End If
          If BFComunicator.ItemValue <> "" Then
            fAnalogWritings(ndx) = CSng(BFComunicator.ItemValue)
            iAnalogWritings(ndx) = CInt(fAnalogWritings(ndx))
          End If
        Next ndx
        
      ElseIf tipo_misura = DIGITALI Then
        
        For ndx = 0 To iMaxDO
          If Val(LineaBFlab) > 0 Then
            BFComunicator.CurrentItem = LineaBFlab & " " & "DO" & ndx
          Else
            BFComunicator.CurrentItem = "DO" & ndx
          End If
          DigitalWritings(ndx) = Val(BFComunicator.ItemValue)
        Next ndx
        
      End If
    End If
    
    Exit Sub
    
GestErr:
    warn_msg = "LeggiBFComunicator: " & Err.Description
    Err.Clear
    Resume Next

End Sub

