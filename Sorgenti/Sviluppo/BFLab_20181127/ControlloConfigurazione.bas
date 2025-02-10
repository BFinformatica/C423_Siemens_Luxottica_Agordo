Attribute VB_Name = "ControllaConfigurazione"
Option Explicit

'luca 05/10/2016
Sub ControlloConfigurazione()
    
    Dim iIDParametro As Integer
    
    On Error GoTo Gesterrore
    
    For iIDParametro = 0 To gnNroParametriStrumenti
        'modifica delle soglie di attenzione istantanee/orarie e giornaliere
        If ParametriStrumenti(iIDParametro).AttivaControlloConfigurazioneSoglie Then Call ControlloConfigurazioneMisureSoglie(iIDParametro)
        
        'modifica dei parametri di QAL2 / QAL3
        If ParametriStrumenti(iIDParametro).AttivaControlloConfigurazioneQAL2QAL3 Then Call ControlloConfigurazioneMisureQAL2QAL3(iIDParametro)
                
        'modifica valore stimato
        If ParametriStrumenti(iIDParametro).AttivaControlloConfigurazioneValoreStimato Then Call ControlloConfigurazioneMisureValoreStimato(iIDParametro)
    Next iIDParametro
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloConfigurazione " + Error(Err), 1, OPC)

End Sub

'Federica dicembre 2017 - Diviso controllo e aggiornamento soglie per tipologia
Private Sub ControlloConfigurazioneMisureSoglie(IndiceParametro As Integer)
    
    Dim SAttenzione As Double
    Dim SAllarme As Double
    Dim CambiatoQualcosa As Boolean
    Dim strInizioTag As String
    
    On Error GoTo Gesterrore
    
    'leggo le tag Wincc e le salvo su delle variabili
    'luca 05/05/2016 accrocchio perchè se faccio la conversione diretta in double da LeggiTag non converte correttamente
    
    CambiatoQualcosa = False
    strInizioTag = "CONFIG" & CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(IndiceParametro).CodiceParametro, "000")
    
    'Orarie
    SAttenzione = CDbl(CStr(LeggiTag(strInizioTag & "_SATT")))
    SAllarme = CDbl(CStr(LeggiTag(strInizioTag & "_SALL")))
    If (ParametriStrumenti(IndiceParametro).SogliaAttenzione <> SAttenzione) Or (ParametriStrumenti(IndiceParametro).SogliaAllarme <> SAllarme) Then
        Call UpdateParCalSoglie(IndiceParametro, "C11", SAttenzione)
        Call UpdateParCalSoglie(IndiceParametro, "C12", SAllarme)
        CambiatoQualcosa = True
    End If

    'Giornaliere
    SAttenzione = CDbl(CStr(LeggiTag(strInizioTag & "_SATT_GIORNO")))
    SAllarme = CDbl(CStr(LeggiTag(strInizioTag & "_SALL_GIORNO")))
    If (ParametriStrumenti(IndiceParametro).SogliaAttenzioneGiornaliera <> SAttenzione) Or (ParametriStrumenti(IndiceParametro).SogliaAllarmeGiornaliera <> SAllarme) Then
        Call UpdateParCalSoglie(IndiceParametro, "C75", SAttenzione)
        Call UpdateParCalSoglie(IndiceParametro, "C76", SAllarme)
        CambiatoQualcosa = True
    End If
    
    'Mensili
    SAttenzione = CDbl(CStr(LeggiTag(strInizioTag & "_SATT_MESE")))
    SAllarme = CDbl(CStr(LeggiTag(strInizioTag & "_SALL_MESE")))
    If (ParametriStrumenti(IndiceParametro).SogliaAttenzioneMensile <> SAttenzione) Or (ParametriStrumenti(IndiceParametro).SogliaAllarmeMensile <> SAllarme) Then
        Call UpdateParCalSoglie(IndiceParametro, "L10", SAttenzione)
        Call UpdateParCalSoglie(IndiceParametro, "L11", SAllarme)
        CambiatoQualcosa = True
    End If
    
    If CambiatoQualcosa Then
        'ricarico configurazione
        ScriviTag CStr(NumeroLinea) & "_LeggiConfig", 0
        Call WindasLog("ControlloConfigurazioneMisureSoglie: aggiornata configurazione soglie parametro " & ParametriStrumenti(IndiceParametro).NomeParametro, 0, OPC)
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloConfigurazioneMisureSoglie " + Error(Err), 1, OPC)

End Sub

''Federica novembre 2017 - Aggiunta gestione soglie mensili
'Private Sub UpdateParCalSoglie(ByVal IndiceParametro As Integer, ByVal SAttenzione As Double, ByVal SAllarme As Double, _
'    ByVal SAttenzioneGiornaliera As Double, ByVal SAllarmeGiornaliera As Double, SAttenzioneMensile As Double, SAllarmeMensile As Double)
'
'    Dim queryOK(5) As Boolean
'    Dim Utente As String
'    Dim i As Integer
'
'    On Error GoTo GestErrore
'
'    #If versione = 2 Then
'        Utente = LeggiTag("UtenteCorrente")
'    #Else
'        Utente = LeggiTag("@CurrentUser")
'    #End If
'
'    queryOK(0) = AggiornaConfigurazione("WAS_MEASURES", "C11", Utente, CStr(SAttenzione), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'    queryOK(1) = AggiornaConfigurazione("WAS_MEASURES", "C12", Utente, CStr(SAllarme), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'    queryOK(2) = AggiornaConfigurazione("WAS_MEASURES", "C75", Utente, CStr(SAttenzioneGiornaliera), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'    queryOK(3) = AggiornaConfigurazione("WAS_MEASURES", "C76", Utente, CStr(SAllarmeGiornaliera), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'    queryOK(4) = AggiornaConfigurazione("WAS_MEASURES", "L10", Utente, CStr(SAttenzioneMensile), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'    queryOK(5) = AggiornaConfigurazione("WAS_MEASURES", "L11", Utente, CStr(SAllarmeMensile), "C2='" & CStr(NomeParametro(IndiceParametro)) & "' AND CM_StationCode = '" & CStr(NumeroLinea) & "_SiCEMS'", NumeroLinea)
'
'    'luca 28/04/2016 se la funzione di windasfwk .AggiornaConfigurazione ritorna false significa che è andata in errore
'    For i = 0 To 5
'        If Not queryOK(i) Then
'            Call WindasLog("UpdateParCalSoglie: query aggiornamento configurazione soglie n° " & CStr(i) & " per il parametro " & CStr(NomeParametro(IndiceParametro)) & " non eseguita correttamente", 0, OPC)
'        End If
'    Next i
'
'    Exit Sub
'
'GestErrore:
'    Call WindasLog("UpdateParCalSoglie " + Error(Err), 1, OPC)
'
'End Sub

Private Sub UpdateParCalSoglie(ByVal IndiceParametro As Integer, ByVal Campo As String, ByVal valore As Double)
    
    Dim queryOK As Boolean
    Dim Utente As String
    Dim i As Integer
    
    On Error GoTo Gesterrore
    
    #If versione = 2 Then
        Utente = LeggiTag("UtenteCorrente")
    #Else
        Utente = LeggiTag("@CurrentUser")
    #End If
    
    queryOK = AggiornaConfigurazione("WAS_MEASURES", Campo, Utente, CStr(valore), "C2='" & ParametriStrumenti(IndiceParametro).NomeParametro & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)

    'luca 28/04/2016 se la funzione di windasfwk .AggiornaConfigurazione ritorna false significa che è andata in errore
    If Not queryOK Then
        Call WindasLog("UpdateParCalSoglie: query aggiornamento configurazione soglie campo " & Campo & " per il parametro " & ParametriStrumenti(IndiceParametro).NomeParametro & " non eseguita correttamente", 0, OPC)
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("UpdateParCalSoglie " + Error(Err), 1, OPC)

End Sub

Private Sub ControlloConfigurazioneMisureQAL2QAL3(IndiceParametro As Integer, Optional MCablato As Variant, Optional QCablato As Variant, Optional RangeCablato As Variant, Optional ICCablato As Variant, Optional DataQAL2Cablato As Variant, Optional ZeroRefCablato As Variant, Optional SpanRefCablato As Variant)
    
    Dim ModificataConfigurazione As Boolean
    Dim ParametroM As Double
    Dim ParametroQ As Double
    Dim ParametroRange As Double
    Dim ParametroIC As Double
    Dim ParametroDataQAL2 As String
    Dim ParametroZeroRef As Double
    Dim ParametroSpanRef As Double
    Dim DataQAL2Temp As String
    Dim strInizioTag As String
    Dim tmp As String
    
    On Error GoTo Gesterrore
    
    ModificataConfigurazione = False
    strInizioTag = CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(IndiceParametro).CodiceParametro, "000")
    
    'leggo le tag Wincc e le salvo su delle variabili (controllando se vi sono eventuali valori forzati opzionali)
    If Not IsMissing(MCablato) Then
        ParametroM = CDbl(MCablato)
    Else
        'luca 03/05/2016 accrocchio perchè se faccio la conversione diretta in double da LeggiTag non converte correttamente
        ParametroM = CDbl(CStr(LeggiTag(strInizioTag & "_QAL2_M")))
    End If
    
    If Not IsMissing(QCablato) Then
        ParametroQ = CDbl(QCablato)
    Else
        ParametroQ = CDbl(CStr(LeggiTag(strInizioTag & "_QAL2_Q")))
    End If
    
    If Not IsMissing(RangeCablato) Then
        ParametroRange = CDbl(RangeCablato)
    Else
        ParametroRange = CDbl(CStr(LeggiTag(strInizioTag & "_QAL2_RANGE")))
    End If
    
    If Not IsMissing(ICCablato) Then
        ParametroIC = CDbl(ICCablato)
    Else
        ParametroIC = CDbl(CStr(LeggiTag(strInizioTag & "_QAL2_IC")))
    End If
    
    If Not IsMissing(DataQAL2Cablato) Then
        ParametroDataQAL2 = CStr(DataQAL2Cablato)
    Else
        ParametroDataQAL2 = CStr(LeggiTag(strInizioTag & "_QAL2_DATE"))
    End If
    
    If Not IsMissing(ZeroRefCablato) Then
        ParametroZeroRef = CDbl(ZeroRefCablato)
    Else
        ParametroZeroRef = CDbl(CStr(LeggiTag(strInizioTag & "_QAL3_ZEROREF")))
    End If
    
    If Not IsMissing(SpanRefCablato) Then
        ParametroSpanRef = CDbl(SpanRefCablato)
    Else
        ParametroSpanRef = CDbl(CStr(LeggiTag(strInizioTag & "_QAL3_SPANREF")))
    End If
    
    'luca 28/04/2016 controllo se la data è vuota o con un valore corretto
    tmp = ParametriStrumenti(IndiceParametro).DataQAL2
    If tmp <> "" Then
        DataQAL2Temp = Right(tmp, 2) & "/" & Mid(tmp, 5, 2) & "/" & Left(tmp, 4)
    Else
        DataQAL2Temp = ""
    End If
    
    'se vi è anche solo un parametro diverso da quello configurato salvo nel DB
    If ParametriStrumenti(IndiceParametro).m <> ParametroM _
    Or ParametriStrumenti(IndiceParametro).q <> ParametroQ _
    Or ParametriStrumenti(IndiceParametro).Range <> ParametroRange _
    Or ParametriStrumenti(IndiceParametro).IntervalloConfidenza <> ParametroIC _
    Or DataQAL2Temp <> ParametroDataQAL2 _
    Or ParametriStrumenti(IndiceParametro).ZeroTeorico <> ParametroZeroRef _
    Or ParametriStrumenti(IndiceParametro).SpanTeorico <> ParametroSpanRef Then
        ModificataConfigurazione = True
    End If
    
    If ModificataConfigurazione Then
        
        'luca 25/07/2016 se NOx salvo zero e span di riferimento anche per NO
        If IndiceParametro = IngressoNOX Then
            Call UpdateParCal(IndiceParametro, ParametroM, ParametroQ, ParametroRange, ParametroIC, ParametroDataQAL2, ParametroZeroRef, ParametroSpanRef)
            Call UpdateParCal(IngressoNO, 0, 0, 0, 0, "", ParametroZeroRef, ParametroSpanRef)
        Else
            Call UpdateParCal(IndiceParametro, ParametroM, ParametroQ, ParametroRange, ParametroIC, ParametroDataQAL2, ParametroZeroRef, ParametroSpanRef)
        End If
        'ricarico configurazione
        ScriviTag CStr(NumeroLinea) & "_LeggiConfig", 0
        
        Call WindasLog("ControlloConfigurazioneMisureQAL2QAL3: aggiornata configurazione misure parametro " & ParametriStrumenti(IndiceParametro).NomeParametro, 0, OPC)

    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloConfigurazioneMisureQAL2QAL3 " + Error(Err), 1, OPC)

End Sub

'luca 28/04/2016 revisiono updateparcal
Public Sub UpdateParCal(ByVal IndiceParametro As Integer, ByVal m As Double, ByVal q As Double, ByVal Range As Double, ByVal IC As Double, ByVal DataQAL2 As String, ByVal ZeroRef As Double, ByVal SpanRef As Double)
    
    Dim queryOK(6) As Boolean
    Dim Utente As String
    Dim DataQAL2Temp As String
    Dim NomePar As String
    
    On Error GoTo Gesterrore
    
    #If versione = 2 Then
        Utente = LeggiTag("UtenteCorrente")
    #ElseIf versione = 3 Then
        Utente = "Amministratore"
    #Else
        Utente = LeggiTag("@CurrentUser")
    #End If
    
    If DataQAL2 <> "" Then
        DataQAL2Temp = Format(CDate(DataQAL2), "yyyymmdd")
    Else
        DataQAL2Temp = ""
    End If
    
    NomePar = ParametriStrumenti(IndiceParametro).NomeParametro
    queryOK(0) = AggiornaConfigurazione("WAS_MEASURES", "C41", Utente, CStr(m), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(1) = AggiornaConfigurazione("WAS_MEASURES", "C42", Utente, CStr(q), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(2) = AggiornaConfigurazione("WAS_MEASURES", "C43", Utente, CStr(Range), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(3) = AggiornaConfigurazione("WAS_MEASURES", "C44", Utente, CStr(IC), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(4) = AggiornaConfigurazione("WAS_MEASURES", "C45", Utente, CStr(DataQAL2Temp), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(5) = AggiornaConfigurazione("WAS_MEASURES", "C58", Utente, CStr(ZeroRef), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)
    queryOK(6) = AggiornaConfigurazione("WAS_MEASURES", "C59", Utente, CStr(SpanRef), "C2='" & NomePar & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)

    'luca 28/04/2016 se la funzione di windasfwk .AggiornaConfigurazione ritorna false significa che è andata in errore
    Dim i As Integer
    For i = 0 To 6
        If Not queryOK(i) Then
            Call WindasLog("UpdateParCal: query aggiornamento configurazione misure n° " & CStr(i) & " per il parametro " & NomePar & " non eseguita correttamente", 0, OPC)
        End If
    Next i
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("UpdateParCal " + Error(Err), 1, OPC)

End Sub

'luca 05/05/2016
Private Sub ControlloConfigurazioneMisureValoreStimato(IndiceParametro As Integer)

    Dim VStimato As Double
    
    On Error GoTo Gesterrore
    
    'luca 05/05/2016 controllo che il parametro selezionato sia impostato come parametro calcolato
    If ParametriStrumenti(IndiceParametro).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        'leggo le tag Wincc e le salvo su delle variabili
        'luca 05/05/2016 accrocchio perchè se faccio la conversione diretta in double da LeggiTag non converte correttamente
        VStimato = CDbl(CStr(LeggiTag("CONFIG" & CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(IndiceParametro).CodiceParametro, "000") & "_VALSTIMATO")))
        
        'se vi è almeno un parametro diverso da quello configurato salvo nel DB
        If TrasformaInDbl(ParametriStrumenti(IndiceParametro).OpzioniAcquisizione) <> VStimato Then
            Call UpdateParCalValStimato(IndiceParametro, VStimato)
            'ricarico configurazione
            ScriviTag CStr(NumeroLinea) & "_LeggiConfig", 0
            
            Call WindasLog("ControlloConfigurazioneMisureValoreStimato: aggiornata configurazione valore stimato " & ParametriStrumenti(IndiceParametro).NomeParametro, 0, OPC)
        End If
    Else
        Call WindasLog("ControlloConfigurazioneMisureValoreStimato: ATTENZIONE CONFIGURAZIONE PARAMETRO " & ParametriStrumenti(IndiceParametro).NomeParametro & " NON CORRETTA", 0, OPC)
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloConfigurazioneMisureValoreStimato " + Error(Err), 1, OPC)

End Sub

'luca 05/10/2016 updateparcal per il solo valore stimato
Private Sub UpdateParCalValStimato(ByVal IndiceParametro As Integer, ByVal VStimato As Double)
    
    Dim queryOK As Boolean
    Dim Utente As String
    
    On Error GoTo Gesterrore
    
    #If versione = 2 Then
        Utente = LeggiTag("UtenteCorrente")
    #Else
        Utente = LeggiTag("@CurrentUser")
    #End If
    
    queryOK = AggiornaConfigurazione("WAS_MEASURES", "C19", Utente, CStr(VStimato), "C2='" & ParametriStrumenti(IndiceParametro).NomeParametro & "' AND CM_StationCode = '" & gsClienteDi & "'", NumeroLinea)

    'luca 28/04/2016 se la funzione di windasfwk .AggiornaConfigurazione ritorna false significa che è andata in errore
    If Not queryOK Then
        Call WindasLog("UpdateParCalValStimato: query aggiornamento configurazione valore stimato per il parametro " & ParametriStrumenti(IndiceParametro).NomeParametro & " non eseguita correttamente", 0, OPC)
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("UpdateParCalValStimato " + Error(Err), 1, OPC)

End Sub
