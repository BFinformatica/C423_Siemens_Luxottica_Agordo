Attribute VB_Name = "WinCC"
Option Explicit
Global VersioneTag As String
Global mcp
'luca 22/07/2016
Global NumeroLineaBFData As Integer
Global Comunicator As Object

Function LeggiTag(NomeTag)

    On Error GoTo GestErrore

    'Alby Novembre 2013
    'Routine che gestisce le tag
    'sia in modalità WinCC script che VB6 OPC
        
    'Versione WinCC ActiveX
    'luca maggio 2018
    #If versione = 2 Then
        'Alby Marzo 2018
        Comunicator.CurrentItem = NomeTag
        LeggiTag = Comunicator.ItemValue
    #Else
        LeggiTag = OPC.LeggiTagWinCC(NomeTag)
    #End If

    Exit Function
    
GestErrore:
    Call WindasLog("LeggiTag " + Error(Err), 1)

End Function

Sub ScriviTag(NomeTag, Variabile)
    
    'luca maggio 2018
    #If versione = 2 Then
        'Alby Marzo 2018
        On Error Resume Next
        
        Comunicator.AddItem NomeTag
        Comunicator.CurrentItem = NomeTag
        Comunicator.ItemValue = Variabile
    #Else
        Call ScriviTagWinCC(NomeTag, Variabile)
    #End If
    
    Exit Sub
        
    
    Exit Sub

GestErrore:
    Call WindasLog("ScriviTag ", 1)
    
End Sub

Sub InizializzaWinCC()

    On Error GoTo GestErrore

    'Alby Giugno 2016
    #If versione = 2 Then
        Call InizializzaComunicator
    #Else
        If VersioneTag = "OPC" Then
            Call WindasLog("BFdata inizializzo connessione a WinCC " + VersioneTag, 0)
            Call OPC.InizializzaOPC
        Else
            Set mcp = Nothing
            Set mcp = CreateObject("WinCC-Runtime-Project")
        End If
    #End If
 
  
    Exit Sub
    
GestErrore:
    Call WindasLog("InizializzaWinCC " + Error(Err), 1)
    

End Sub
Function LeggiTagWinCC(NomeTag)

    Dim ErroreConnessione

    On Error GoTo GestErrore

Riprova:
    
    'Alby Giugno 2016
    If VersioneTag = "OPC" Then
        LeggiTagWinCC = OPC.LeggiTagOPC(NomeTag)
        Exit Function
    Else
        LeggiTagWinCC = mcp.GetValue(NomeTag)
    End If
    
    'Alby Febbraio 2014 se manca WinCC
    If IsEmpty(LeggiTagWinCC) Then
        'End

ErroreWinCC:
                
        ErroreConnessione = 1
        Call WindasLog("Attenzione WinCC non operativo! ", 1)
        Call InizializzaWinCC
        Call Ritardo(10)
    
        
        GoTo Riprova
    Else
        If ErroreConnessione = 1 Then
            Call WindasLog("Connessione con WinCC ripristinata", 0)
        End If
    End If
    
    Exit Function
    
GestErrore:
    If Err = 462 Then
        Call WindasLog("Attenzione WinCC non operativo! ", 1)
        GoTo ErroreWinCC
    Else
        Call WindasLog("LeggiTagWinCC " + Error(Err), 1)
    End If
    Resume Next
    
End Function

'luca maggio 2018
Sub InizializzaComunicator()
           
    'Alby Marzo 2018
    Set Comunicator = GetObject("", "BFcomunicator.cloggerdata")
    Comunicator.startdatasharing
    
    Exit Sub
    
GestErrore:
    Call WindasLog("InizializzaComunicator errore: " + Error(Err), 1)

End Sub

Sub ScriviTagWinCC(NomeTag, Variabile)

    Dim StrNomeTag As String

    On Error GoTo GestErrore
    
    'Alby Giugno 2016
    If VersioneTag = "OPC" Then
        Call OPC.ScriviTagOPC(NomeTag, Variabile)
        Exit Sub
    Else
        mcp.setvalue NomeTag, Variabile
    End If
    
    Exit Sub
    
GestErrore:
    If Err = 462 Then
        Call WindasLog("WinCC non operativo! ", 1)
    Else
        Call WindasLog("ScriviTagWinCC " + Error(Err), 1)
    End If

End Sub

