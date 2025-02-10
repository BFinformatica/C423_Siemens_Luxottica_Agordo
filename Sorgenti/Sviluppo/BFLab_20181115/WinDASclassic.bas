Attribute VB_Name = "WinDASclassic"
Option Explicit
Global Comunicator As Object

Sub InizializzaComunicator()
           
    'Alby Marzo 2018
    Set Comunicator = GetObject("", "BFcomunicator.cloggerdata")
    Comunicator.startdatasharing
    
    Exit Sub
    
GestErrore:
    Call WindasLog("InizializzaComunicator errore: " + Error(Err), 1, OPC)

End Sub
