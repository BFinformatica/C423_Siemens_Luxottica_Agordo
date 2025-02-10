Attribute VB_Name = "INIFile"
'*************************************************************************************
'* INIFile.BAS - Funzioni di Lettura Del file INI
'*
'*    Copyright (c) 2010 SIEI Automazioni  s.r.l.
'*
'*    Dichiarazioni e Funzioni Generali Gestione Interfacciamento
'*    Driver PLC (S7 SoftNet)
'*
'*----------------------------------------------------------
'*
'*Progetto           : COMUNICAZIONI S7 SOFTNET
'*Release            : 2.00
'*Sviluppo           :
'*    Piattaforma    : WIN XP
'*    Ling. & Comp.  : Visual Basic 6.00
'*Piattaforma Target : WIN2000,WINXP,WIN7
'*Data Creazione     : 03/11/10
'*Data Aggiornamento : 03/11/10
'*Autore             : Breda Luca
'*
'*----------------------------------------------------------
'*History:
'*  Data    Autore    Note
'*----------------------------------------------------------
'*
'*************************************************************************************
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long




'**************************************************************
'* Funzione     : Tokenize(szBuffer As String, szSeparators As String, rgszArguments() As String) As Integer
'* Scopo        : Estrazione di Token da una stringa
'* INGRESSI     :
'*      szBuffer As String : stringa di riferimento
'*      szSeparators As String : lista dei separatori di token
'*      rgszArguments() As String : Vettore dei token
'*      Optional bNullStrings As Variant: se TRUE verifica e testa campo null
'* USCITE       :
'*      Numero di elementi estratti
'**************************************************************
Function Tokenize(szBuffer As String, szSeparators As String, rgszArguments() As String, Optional bNullStrings As Variant) As Integer
    
    Dim iArgsAmount As Integer
    Dim iIndex      As Integer
    Dim iStart      As Integer
    Dim iFound      As Integer
    
    On Local Error GoTo TokenError
    
    If (IsMissing(bNullStrings)) Then bNullStrings = False
    iArgsAmount = 0
    iStart = 1
    
    'Ciclo di tokenizzazione
    For iIndex = 1 To Len(szBuffer)
        iFound = InStr(szSeparators, Mid$(szBuffer, iIndex, 1))
        If (iFound > 0) Then
            rgszArguments(iArgsAmount) = Mid$(szBuffer, iStart, iIndex - iStart)
    
            'Copia l' argomento
            If (bNullStrings) Or (Len(rgszArguments(iArgsAmount))) Then
                iArgsAmount = iArgsAmount + 1
            End If
            'Punta al prossimo elemento
            iStart = iIndex + 1
        End If
    Next
    'Copia l' ultimo argomento
    If (iStart < iIndex) Then
        rgszArguments(iArgsAmount) = Mid$(szBuffer, iStart, Len(szBuffer) - iStart + 1)
        If (bNullStrings) Or (Len(rgszArguments(iArgsAmount))) Then
            iArgsAmount = iArgsAmount + 1
        End If
    End If
    
    'Splitting
    Tokenize = iArgsAmount
    Exit Function
    
TokenError:
    Resume Next
    
End Function



'**********************************************************************
'* Class        : LabelPrinter
'* Metodo       : ReadFileIni()
'* Scopo        : Legge dal file INI la configurazione
'* INGRESSI     :
'*      NESSUNO
'* USCITE       :
'*      NESSUNO
'**********************************************************************
Public Function ReadFileIni()
Dim szFileINI     As String
Dim szIPAddress   As String
Dim iArgs         As Integer
Dim rgszArgs(10)  As String
Dim iRetCode      As Long

On Local Error Resume Next
    
    ReadFileIni = True
    'Nome del file ini
    szFileINI = App.Path + "\S7TCPIP.INI"
    With rgPLCDef(ID_MASTER)
        .szDescrizione = "MASTER"
        szIPAddress = "111.111.111.111        "
        iRetCode = GetPrivateProfileString("MASTER", "IPAddress", "192.168.2.10", szIPAddress, Len(szIPAddress), szFileINI)
        szIPAddress = VBA.Left$(szIPAddress, iRetCode)
        
        '**** michele - ping
        IP_Master = szIPAddress

        iArgs = Tokenize(szIPAddress, ".", rgszArgs)
        If iArgs = 4 Then
          .iIPAddress(1) = Val(rgszArgs(0))
          .iIPAddress(2) = Val(rgszArgs(1))
          .iIPAddress(3) = Val(rgszArgs(2))
          .iIPAddress(4) = Val(rgszArgs(3))
          .bInUso = True
        Else
            ReadFileIni = False
        End If
        
        szIPAddress = "100       "
        iRetCode = GetPrivateProfileString("MASTER", "DB", "100", szIPAddress, Len(szIPAddress), szFileINI)
        .DB = Val(VBA.Left$(szIPAddress, iRetCode))

    End With
    With rgPLCDef(ID_SLAVE)
        .szDescrizione = "SLAVE"
        szIPAddress = "111.111.111.111        "
        iRetCode = GetPrivateProfileString("SLAVE", "IPAddress", "192.168.2.11", szIPAddress, Len(szIPAddress), szFileINI)
        szIPAddress = VBA.Left$(szIPAddress, iRetCode)
        iArgs = Tokenize(szIPAddress, ".", rgszArgs)
        If iArgs = 4 Then
          .iIPAddress(1) = Val(rgszArgs(0))
          .iIPAddress(2) = Val(rgszArgs(1))
          .iIPAddress(3) = Val(rgszArgs(2))
          .iIPAddress(4) = Val(rgszArgs(3))
          .bInUso = True
          
        Else
            ReadFileIni = False
        End If
        
        szIPAddress = "100       "
        iRetCode = GetPrivateProfileString("SLAVE", "DB", "100", szIPAddress, Len(szIPAddress), szFileINI)
        .DB = Val(VBA.Left$(szIPAddress, iRetCode))

    End With
End Function

