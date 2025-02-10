Attribute VB_Name = "DEMO_PLC"
'*************************************************************************************
'* DEMO.BAS - Dichiarazioni delle Funzioni di Interfacciamento col PLC
'*
'*    Copyright (c) 2010 SIEI Automazioni s.r.l.
'*
'*    Dichiarazioni e Funzioni Generali Gestione Interfacciamento
'*
'*----------------------------------------------------------
'*
'*Progetto           : COMUNICAZIONI S7 SOFTNET
'*Release            : 2.00
'*Sviluppo           :
'*    Piattaforma    : WIN XP
'*    Ling. & Comp.  : Visual Basic 6.00
'*Piattaforma Target : WIN2000,WINXP,WIN7
'*Data Creazione     : 29/10/10
'*Data Aggiornamento : 29/10/10
'*Autore             : Breda Luca
'*
'*----------------------------------------------------------
'*History:
'*  Data    Autore    Note
'*----------------------------------------------------------
'*
'*************************************************************************************

Type PLCDATA_DEF
   ValoreReale1     As Single
   ValoreReale2     As Single
   ValoreReale3     As Single
   ValoreReale4     As Single
End Type

    '***********************************************************************************
    '***********************************************************************************
    '****                                                                           ****
    '****                D E F I N I Z I O N E   C O S T A N T I                    ****
    '****                                                                           ****
    '***********************************************************************************
    '***********************************************************************************

    'ID dei PLC
Public Const ID_MASTER = 1
Public Const ID_SLAVE = 2
Public Const MAX_PLC = 2

Public Const ID_MASTER_1 = 1
Public Const ID_MASTER_2 = 2
Public Const ID_MASTER_3 = 3
Public Const ID_MASTER_4 = 4
Public Const ID_MASTER_5 = 5

Public Const PLC_CPU_TYPE_S7_400 = 3
Public Const PLC_CPU_SLOT_S7_400 = 3
Public Const PLC_CPU_TYPE_S7_300 = 2
Public Const PLC_CPU_SLOT_S7_300 = 2



Public Const MAX_CYCLE_PLC_ERR = 5
        

    '***********************************************************************************
    '***********************************************************************************
    '****                                                                           ****
    '****          D E F I N I Z I O N E   S T R U T T U R E   D A T I              ****
    '****                                                                           ****
    '***********************************************************************************
    '***********************************************************************************
        
    'Definizione della Struttura di Definizione per il PLC
Public Type PLCDEF_DEFINIZIONE
        'Configurazione
    szDescrizione                   As String
    iTipo                           As Integer              'Tipo di PLC
    iIPAddress(1 To 4)              As Integer              'Indirizzo IP
        'Run-Time
    bInUso                          As Boolean
    hPLCHandle                      As Long                 'Handle della Connessione Fisica S7 Softnet
    bOnLine                         As Boolean              'Connessione con PLC attiva
    lComErrors                      As Long                 'Errori di comunicazione
End Type

        
        
    '***********************************************************************************
    '***********************************************************************************
    '****                                                                           ****
    '****           A L L O C A Z I O N E   D A T I   C O N D I V I S E             ****
    '****                                                                           ****
    '***********************************************************************************
    '***********************************************************************************
 
'STRUTTURA DI DEFINIZIONE DEL PLC
'Public rgPLCDef(ID_MASTER To ID_SLAVE)      As PLCDEF_DEFINIZIONE            'Struttura di definzione dei PLC
Public rgPLCDef(ID_MASTER_1 To ID_MASTER_5)      As PLCDEF_DEFINIZIONE            'Struttura di definzione dei PLC

Public Function DisconnectPLC(iIDPLC As Integer) As Boolean
    Dim lRetCode        As Long

    On Local Error Resume Next
    
        'Setta il codice di ritorno
    DisconnectPLC = False
    
    With rgPLCDef(iIDPLC)
    
            'Controlla se ha superato i cicli per disconnetterlo
        If (.lComErrors >= MAX_CYCLE_PLC_ERR) Then
                         
                'Se e'definito e connesso
            If (.bInUso) And (.hPLCHandle <> INVALID_HANDLE) Then
               lRetCode = PLCClose(.hPLCHandle)
               .hPLCHandle = INVALID_HANDLE
               .bOnLine = False
                    'Setta il codice di ritorno
                DisconnectPLC = True
            End If
        Else
            .bOnLine = False
        End If
    End With
End Function


Public Function ConnectPLC(iIDPLC As Integer, Optional ulTimeOut As Long = PLC_DEF_TIMEOUT) As Boolean
    
    Dim lRetCode As Long
    Static NumTentativi As Integer

    On Local Error GoTo GestErrore
    
    NumTentativi = 0
    
    '***** Setta il codice di ritorno *****
    ConnectPLC = False
    
       With rgPLCDef(iIDPLC)
                'Se e'definito
            If (.bInUso) Then
                    'Loop di Tentativo di Connessione
                Do
                    DoEvents
                        'Reset del Flag di In Linea
                    .bOnLine = False
                    If (.hPLCHandle = INVALID_HANDLE) Then
                        '*** Apertura Canali di comunicazione
                        '.hPLCHandle = PLCOpenEx(iIDPLC, .iIPAddress(1), .iIPAddress(2), .iIPAddress(3), .iIPAddress(4), PLC_CPU_TYPE_S7_400, PLC_CPU_SLOT_S7_400)
                        .hPLCHandle = PLCOpenEx(iIDPLC, .iIPAddress(1), .iIPAddress(2), .iIPAddress(3), .iIPAddress(4), PLC_CPU_TYPE_S7_300, PLC_CPU_SLOT_S7_300)
                    End If
                        'Se la connessione e' Valida
                    If ((.hPLCHandle <> INVALID_HANDLE) _
                    And (.hPLCHandle <> 0)) Then
                            'OK: Il PLC è connesso
                        .bOnLine = True
                        ConnectPLC = True
                    End If
                    
                    If NumTentativi = 0 Then
                        Exit Do
                    Else
                        NumTentativi = NumTentativi + 1
                    End If
                    
                Loop While (.bOnLine = False)
            End If
        End With
        Exit Function
        
GestErrore:
    Debug.Print Now & " ConnectPLC: " & Error(Err)

End Function




Public Function ConnectAll() As Boolean
    Dim iIDPLC          As Integer
    Dim lRetCode        As Long
    
    On Local Error Resume Next
   
        
        'Cicla per tutti i PLC
    For iIDPLC = 1 To MAX_PLC

    
    Next
        'Setta il codice di ritorno
    ConnectAll = True
End Function

Public Function DisconnectAll() As Boolean
    Dim iIDPLC          As Integer
    Dim lRetCode        As Long

    On Local Error Resume Next
    
        'Cicla per tutti i PLC
    For iIDPLC = CurrentPLC To MAX_PLC
                DisconnectPLC (iIDPLC)
    Next
        'Setta il codice di ritorno
    DisconnectAll = True
End Function

