Attribute VB_Name = "S7TCPIP_PLC"
'*************************************************************************************
'* S7TCPIP.BAS - Dichiarazioni delle Funzioni Importate da DLL
'*
'*    Copyright (c) 2009 SIEI Automazioni  s.r.l.
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
'*Data Creazione     : 21/08/09
'*Data Aggiornamento : 21/08/09
'*Autore             : Breda Luca
'*
'*----------------------------------------------------------
'*History:
'*  Data    Autore    Note
'*----------------------------------------------------------
'*
'*************************************************************************************

Option Explicit


    ' Dichiarazioni
Public Const PLC_DEF_TIMEOUT = 2000
Public Const PLC_ERROR = 0

Public Const INVALID_HANDLE = 0
Public Const DEMO_HANDLE = 1

    '**************************************************************************
    '******  Prototipi delle Funzioni Importate di gestione S7 SOFTNET  *******
    '**************************************************************************
    '*** Gestione Connessione
Public Declare Function PLCOpen Lib "S7TCPIP.DLL" (ByVal iPLCID As Integer, ByVal iPLCAddress As Integer) As Long
Public Declare Function PLCOpenEx Lib "S7TCPIP.DLL" (ByVal iPLCID As Integer, ByVal iPLCAddress1 As Integer, ByVal iPLCAddress2 As Integer, ByVal iPLCAddress3 As Integer, ByVal iPLCAddress4 As Integer, ByVal iPLCType As Integer, ByVal iCPUSlot As Integer) As Long
Public Declare Function PLCClose Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long) As Long
Public Declare Function PLCHandle Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long) As Long
    '*** Gestione Letture/Scritture Generiche
Public Declare Function PLCRead Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iType As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
Public Declare Function PLCWrite Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iType As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
    '*** Gestione Letture/Scritture di Bytes su Blocchi DB
Public Declare Function PLCByteRead Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
Public Declare Function PLCByteWrite Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
    '*** Gestione Letture/Scritture di Words su Blocchi DB
Public Declare Function PLCWordRead Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
Public Declare Function PLCWordWrite Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
    '*** Gestione Letture/Scritture di DWords su Blocchi DB
Public Declare Function PLCDWordRead Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
Public Declare Function PLCDWordWrite Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iObjectsAmount As Integer, ByRef pBuffer As Any) As Long
  
    '*** Gestione Bit appoggiati su Blocchi DB
Public Declare Function PLCBitRead Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iBit As Integer, ByVal iBitValue As Integer) As Long
Public Declare Function PLCBitWrite Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iBit As Integer, ByVal iBitValue As Integer) As Long
Public Declare Function PLCBitWait Lib "S7TCPIP.DLL" (ByVal hPLCHandle As Long, ByVal iSegment As Integer, ByVal iStartingObject As Integer, ByVal iBit As Integer, ByVal iBitValue As Integer, ByVal lTimeout As Long) As Long
    '*** Gestione degli Errori
Public Declare Function PLCGetLastError Lib "S7TCPIP.DLL" () As Long
    '*** Funzioni di Setup e Debug
Public Declare Sub PLCSetUp Lib "S7TCPIP.DLL" Alias "PLCSetup" ()
Public Declare Sub PLCDemoMode Lib "S7TCPIP.DLL" (ByVal iDemoMode As Integer)

    '*** Funzioni Ausiliarie per Swapping di Dati
Public Declare Function PLCSwapWord Lib "S7TCPIP.DLL" (ByVal sWord As Integer) As Integer
Public Declare Function PLCSwapDWord Lib "S7TCPIP.DLL" (ByVal sWord As Long) As Long
Public Declare Function PLCSwapLong Lib "S7TCPIP.DLL" (ByVal sWord As Long) As Long
Public Declare Function PLCSwapFloat Lib "S7TCPIP.DLL" (ByVal sWord As Single) As Single
Public Declare Sub PLCSwapBuffer Lib "S7TCPIP.DLL" (ByVal sType As Integer, ByRef pDest As Any, ByRef pSource As Any, ByVal iObjectsAmount As Integer)


    '*** Funzioni Ausiliarie
Public Declare Function ByteBitSet Lib "S7TCPIP.DLL" (ByVal uchValue As Byte, ByVal usBit As Integer, ByVal usBitValue As Integer) As Byte
Public Declare Function WordBitSet Lib "S7TCPIP.DLL" (ByVal usValue As Integer, ByVal usBit As Integer, ByVal usBitValue As Integer) As Integer
Public Declare Function DWordBitSet Lib "S7TCPIP.DLL" (ByVal ulValue As Long, ByVal usBit As Integer, ByVal usBitValue As Integer) As Long
Public Declare Function ByteBitGet Lib "S7TCPIP.DLL" (ByVal uchValue As Byte, ByVal usBit As Integer, ByVal usBitValue As Integer) As Byte
Public Declare Function WordBitGet Lib "S7TCPIP.DLL" (ByVal usValue As Integer, ByVal usBit As Integer, ByVal usBitValue As Integer) As Integer
Public Declare Function DWordBitGet Lib "S7TCPIP.DLL" (ByVal ulValue As Long, ByVal usBit As Integer, ByVal usBitValue As Integer) As Long




    'VARIABILI PUBBLICHE PER GESTIONE COMUNICAZIONI CON PLC
Public bOpened          As Boolean
Public bSegment         As Byte
Public bObjectType      As Byte
Public usFirstObject    As Integer
Public usObjectsAmount  As Integer


