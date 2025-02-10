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
Global Const ID_MASTER = 1
Global Const ID_SLAVE = 2
Global Const MAX_PLC = 2

Global Const PLC_CPU_TYPE_S7_400 = 3
Global Const PLC_CPU_SLOT_S7_400 = 3
Global Const PLC_CPU_TYPE_S7_300 = 2
Global Const PLC_CPU_SLOT_S7_300 = 2
Global Const MAX_CYCLE_PLC_ERR = 2
        

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
    DB                              As Integer              'DB area dati
End Type

        
        
    '***********************************************************************************
    '***********************************************************************************
    '****                                                                           ****
    '****           A L L O C A Z I O N E   D A T I   C O N D I V I S E             ****
    '****                                                                           ****
    '***********************************************************************************
    '***********************************************************************************
 
    'STRUTTURA DI DEFINIZIONE DEL PLC
Public rgPLCDef(ID_MASTER To ID_SLAVE)      As PLCDEF_DEFINIZIONE            'Struttura di definzione dei PLC

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
    
    Dim lRetCode        As Long
    Dim StartTime       As Date

    On Local Error Resume Next
    
    '**** Time Stamp
    StartTime = Now
    
        'Setta il codice di ritorno
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
                    '***** S7 400 *****
                    '.hPLCHandle = PLCOpenEx(iIDPLC, .iIPAddress(1), .iIPAddress(2), .iIPAddress(3), .iIPAddress(4), PLC_CPU_TYPE_S7_400, PLC_CPU_SLOT_S7_400)
                    
                    '***** S7 300 *****
                    '.hPLCHandle = PLCOpenEx(iIDPLC, .iIPAddress(1), .iIPAddress(2), .iIPAddress(3), .iIPAddress(4), PLC_CPU_TYPE_S7_300, PLC_CPU_SLOT_S7_300)
                    
                    'Alby Febbraio 2017 per PLC S7 1200-1500 ci deve essere abilitato in programma PLC get/put e DB non compresse
                    .hPLCHandle = PLCOpenEx(iIDPLC, .iIPAddress(1), .iIPAddress(2), .iIPAddress(3), .iIPAddress(4), PLC_CPU_TYPE_S7_300, 1)
                    
                    ' ConnectPLC = rgPLCDef(iIDPLC).bOnLine
                 End If
                     'Se la connessione e' Valida
                 If ((.hPLCHandle <> INVALID_HANDLE) And (.hPLCHandle <> 0)) Then
                         'OK: Il PLC è connesso
                     .bOnLine = True
                 Else
                 
                     'aspetto un secondo e riprovo
                     Call Ritardo(1)
                     
                 End If
                 
                 If DateDiff("s", StartTime, Now) > 10 Then Exit Do
                 
             Loop While (.bOnLine = False)
         End If
     End With
     ConnectPLC = rgPLCDef(iIDPLC).bOnLine
        
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
    For iIDPLC = ID_PLC1 To MAX_PLC
                DisconnectPLC (iIDPLC)
    Next
        'Setta il codice di ritorno
    DisconnectAll = True
End Function

