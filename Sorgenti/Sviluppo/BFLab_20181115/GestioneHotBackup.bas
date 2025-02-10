Attribute VB_Name = "GestioneHotBackup"
Option Explicit

'Federica giugno 2018
'Da chiamare in "Acquisisce" dopo il controllo del Watchdog
Sub ControlloHotBackup()

    Dim anomalia_partner As Integer
    Dim mio_ruolo As Integer

    On Error GoTo Gesterrore
    
    'Aggiorno il mio Watchdog
    Call AggiornaWatchdogBFLab
    
    If CBool(Generiche(iHotBackup).Par) Then
        anomalia_partner = LeggiTag("ANOMALIA_PARTNER")
        manValoreDigitale(999, 3) = anomalia_partner
    Else
        'Non abilito mai l'allarme
        manValoreDigitale(999, 3) = 0
    End If
    
    'Leggo il mio ruolo dalla tag
    mio_ruolo = LeggiTag("ROLE")
    IsMaster = CBool(mio_ruolo)
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloHotBackup: " & Error(Err()), 1, "OPC")

End Sub

Private Sub AggiornaWatchdogBFLab()
    'Nicolò Gennaio 2018 aggiungo intera sub
On Error GoTo Gesterrore

    Dim nFile As Integer
    nFile = FreeFile
    Open App.Path & "\BFLabWatchdog.txt" For Output As #nFile
    Print #nFile, Format(Now, "yyyymmddhhnnss")
    Close #nFile
    
Exit Sub
Gesterrore:
    Call WindasLog("AggiornaWatchdogBflab: " & Error(Err), 1, "OPC")
End Sub

