Attribute VB_Name = "ElabSalvaDati"
Option Explicit

Sub ElaboraSalvaDatiMedie()

    On Error GoTo GestErrore
    
    'daniele luglio 2013 bolgiano: non dovrebbe servire ciclare per parametro, disabilito?
    For nn = 0 To gnNroParametriStrumenti
        Call ElaboraSalvaDatiAzzeroMatrici
    Next nn

    '**** Rielabora tutte le medie
    For nn = 0 To gnNroParametriStrumenti
    
        'If nn = IngressoIMPIANTO Then Stop
    
        For Ora = 0 To OraFine
        
            'If Ora = 14 Then Stop
            
            For nmedie = 1 To numeromedie(Ora)
            
                '***** reset status monitor per stato prevalente in caso di media < 70% *****
                For sts_n = 0 To 10
                    Stato_Monitor(sts_n, 0) = 0
                    Stato_Monitor(sts_n, 1) = sts_n + 1
                Next sts_n
            
                '**** reset dati *****
                SommatoriaOra = 0
                Call ElaboraSalvaDatiAggiornoContatori
                
                '**** calcola media del periodo
                If ContaOraOK(Ora, nn, 0, nmedie) > 0 Then
                    'luca maggio 2017 mi serve la media corretta dello stato impianto (0 -- 100), per calcolarla devo dividere per 720 e non per contaoraok (contaoraok viene incrementato solo se acquisizione in marcia -> la media sarà sempre pari a 100 o a -9999)
                    If nn = IngressoIMPIANTO Then
                        MedieOra(Ora, nn, 0, nmedie) = SommatoriaOra / 720
                    Else
                        MedieOra(Ora, nn, 0, nmedie) = SommatoriaOra / ContaOraOK(Ora, nn, 0, nmedie)
                    End If
                    If ContaOraOK(Ora, nn, 0, nmedie) / MaxDati >= 0.7 Then
                        StsMedieOra(Ora, nn, 0, nmedie) = "VAL"
                        Call InserisciDecimali(MedieOra(Ora, nn, 0, nmedie), 2)
                    Else
                        StsMedieOra(Ora, nn, 0, nmedie) = StatusMonitor(Stato_Monitor())
                    End If
                Else
                    '***** nessun dato valido *****
                     'luca maggio 2017
                    If nn = IngressoIMPIANTO Then
                        MedieOra(Ora, nn, 0, nmedie) = 0
                        StsMedieOra(Ora, nn, 0, nmedie) = "VAL"
                    Else
                        MedieOra(Ora, nn, 0, nmedie) = -9999
                        StsMedieOra(Ora, nn, 0, nmedie) = StatusMonitor(Stato_Monitor())
                    End If
                End If
                
                'Alby Dicembre 2015 copio dato elaborato da quello tal quale
                'Versalis Priolo BFdata elabora tutte le medie senza eseguire ricalcoli
                MedieOra(Ora, nn, 1, nmedie) = MedieOra(Ora, nn, 0, nmedie)
                StsMedieOra(Ora, nn, 1, nmedie) = StsMedieOra(Ora, nn, 0, nmedie)
                
            Next nmedie
        Next Ora
    Next nn
    
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiMedie " + Error(Err), 1)
    Resume Next
   
End Sub

Private Sub ElaboraSalvaDatiAzzeroMatrici()

    On Error GoTo GestErrore

    '***** reset variabili *****
    For Ora = 0 To 23
        'daniele luglio 2013 bolgiano: gestisco il valore corretto di medie
        'For nmedie = 1 To 60
        For nmedie = 1 To 64
            
            '***** reset stati impianto *****
            PercRegime(Ora, nmedie) = 0
            PercMinTec(Ora, nmedie) = 0
            PercFermo(Ora, nmedie) = 0
            PercSpegnimento(Ora, nmedie) = 0
            PercManutenzione(Ora, nmedie) = 0
            PercGuasto(Ora, nmedie) = 0
            PercAnomalo(Ora, nmedie) = 0
            'luca gennaio 2015 aggiungo reset stato impianto 37
            PercPolveri(Ora, nmedie) = 0
            PercAltro(Ora, nmedie) = 0  'Federica marzo 2018 reset stato impianto 38
            
            '***** reset minimo / massimo grezzi / normalizzati *****
            For xx = 0 To 2
                DoEvents
                ContaTutti_5_secondi(Ora, nn, xx, nmedie) = 0
                ContaOraOK(Ora, nn, xx, nmedie) = 0
                MedieOra(Ora, nn, xx, nmedie) = -9999
                'luca 16/09/2016 lo imposto inizialmente a 0 in quanto non c'è la validità del flusso e pertanto continua a fare la somma anche se portata non presente (sommando valori continui a -9999 prima o dopo va in overflow)
                'DatoFlussoMassa(Ora, nn, xx, nmedie) = -9999
                DatoFlussoMassa(Ora, nn, xx, nmedie) = 0
                minimo(Ora, nn, xx, nmedie) = 999999999
                massimo(Ora, nn, xx, nmedie) = -999999999
                For yy = 0 To 720
                    ValIstPerScarto(Ora, nn, yy, xx) = -9999
                Next yy
                'daniele settembre 2013 bolgiano: azzero contatore apposito per gestione Q_Gas ausiliario
                'daniele settembre 2013 bolgiano: correggo gestione Q_Gas ausiliario
                'intDatiAuxOK(nn) = 0
                intDatiAuxOK(Ora, nn) = 0
            Next xx
        Next nmedie
    Next Ora
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiAzzeroMatrici " + Error(Err), 1)
    Resume fine
fine:

End Sub

Private Sub ElaboraSalvaDatiAggiornoContatori()
    
    'daniele luglio 2013 bolgiano: disabilito dichiarazione (già presente in modulo globale)
    'Dim secondi As Integer
    
    On Error GoTo GestErrore
    
    For secondi = (nmedie - 1) * MaxDati To (MaxDati * nmedie) - 1
        'Alby Dicembre 2015
        'If UCase(SuperTrim(gaConfigurazioneArchivio(nn).STRUM.NomeParametro)) = "IMP_L" & Trim(Str(NumeroLinea)) Then
        If nn = IngressoIMPIANTO Then
            Select Case Valore_5_Secondi(Ora, nn, secondi)
                Case 30
                    '***** in marcia *****
                    PercRegime(Ora, nmedie) = PercRegime(Ora, nmedie) + 1
                    '**** per calcolo media
                    SommatoriaOra = SommatoriaOra + 100
                    ContaOraOK(Ora, nn, 0, nmedie) = ContaOraOK(Ora, nn, 0, nmedie) + 1
                    ContaOraOK(Ora, nn, 1, nmedie) = ContaOraOK(Ora, nn, 0, nmedie)
                    
                Case 31
                    '***** minimo tecnico / accensione *****
                    PercMinTec(Ora, nmedie) = PercMinTec(Ora, nmedie) + 1
                    
                Case 32
                    '***** spegnimento *****
                    PercSpegnimento(Ora, nmedie) = PercSpegnimento(Ora, nmedie) + 1
                    
                Case 33
                    '***** manutenzione *****
                    PercManutenzione(Ora, nmedie) = PercManutenzione(Ora, nmedie) + 1
                    
                Case 34
                    '***** fermo *****
                    PercFermo(Ora, nmedie) = PercFermo(Ora, nmedie) + 1
                    
                Case 35
                    '***** guasto *****
                    PercGuasto(Ora, nmedie) = PercGuasto(Ora, nmedie) + 1
                    
                Case 36
                    '***** anomalo *****
                    PercAnomalo(Ora, nmedie) = PercAnomalo(Ora, nmedie) + 1
                
                'luca gennaio 2015
                Case 37
                    '***** taratura misuratore polveri *****
                    PercPolveri(Ora, nmedie) = PercPolveri(Ora, nmedie) + 1
                    
                Case 38 'Federica marzo 2018
                    '***** Altro Stato Impianto *****
                    PercAltro(Ora, nmedie) = PercAltro(Ora, nmedie) + 1
                    
            End Select
                
            
        Else
                
            '**** per calcolo media
            If Valore_5_Secondi(Ora, nn, secondi) <> -9999 Then

                'luca luglio 2017 gestisco anche il VAH come dato valido
                'If Trim(Status_5_Secondi(Ora, nn, secondi)) = "VAL" Or Trim(Status_5_Secondi(Ora, nn, secondi)) = "AUX" Then
                If InStr("VAL AUX VAH", Trim(Status_5_Secondi(Ora, nn, secondi))) > 0 Then
                    
                    SommatoriaOra = SommatoriaOra + Valore_5_Secondi(Ora, nn, secondi)
                    ContaOraOK(Ora, nn, 0, nmedie) = ContaOraOK(Ora, nn, 0, nmedie) + 1
                    
                    'daniele luglio 2013 bolgiano: gestisco caso AUX su q metano
                    If Trim(Status_5_Secondi(Ora, nn, secondi)) = "AUX" Then
                        intDatiAuxOK(Ora, nn) = intDatiAuxOK(Ora, nn) + 1
                    End If
                    
                    '**** minimo/massimo grezzi *****
                    If Valore_5_Secondi(Ora, nn, secondi) < minimo(Ora, nn, 0, nmedie) Then
                        minimo(Ora, nn, 0, nmedie) = Valore_5_Secondi(Ora, nn, secondi)
                        Call InserisciDecimali(minimo(Ora, nn, 0, nmedie), 2)
                    End If
                        
                    If Valore_5_Secondi(Ora, nn, secondi) > massimo(Ora, nn, 0, nmedie) Then
                        massimo(Ora, nn, 0, nmedie) = Valore_5_Secondi(Ora, nn, secondi)
                        Call InserisciDecimali(massimo(Ora, nn, 0, nmedie), 2)
                    End If
                    
                    ValIstPerScarto(Ora, nn, ContaOraOK(Ora, nn, 0, nmedie), 0) = Valore_5_Secondi(Ora, nn, secondi)
                    
                    '***** normalizzazione dati grezzi = elaborati *****
                    Call ElaboraSalvaDatiNormalizzaIstantaneo(0)
                    
                    'luca luglio 2017 gestisco anche il VAH come dato valido
                    If InStr("VAL VAH", Trim(Status_5_Secondi_N(Ora, nn, secondi))) > 0 Then
                    
                        ContaOraOK(Ora, nn, 1, nmedie) = ContaOraOK(Ora, nn, 1, nmedie) + 1
                        
                        '**** minimo/massimo elaborati *****
                        If Valore_5_Secondi_N(Ora, nn, secondi) < minimo(Ora, nn, 1, nmedie) Then
                            minimo(Ora, nn, 1, nmedie) = Valore_5_Secondi_N(Ora, nn, secondi)
                            Call InserisciDecimali(minimo(Ora, nn, 1, nmedie), 2)
                        End If
                        If Valore_5_Secondi_N(Ora, nn, secondi) > massimo(Ora, nn, 1, nmedie) Then
                            massimo(Ora, nn, 1, nmedie) = Valore_5_Secondi_N(Ora, nn, secondi)
                            Call InserisciDecimali(massimo(Ora, nn, 1, nmedie), 2)
                        End If
                        
                        '**** per StadardDev elaborati *****
                        ValIstPerScarto(Ora, nn, ContaOraOK(Ora, nn, 1, nmedie), 1) = Valore_5_Secondi_N(Ora, nn, secondi)
                    
                    End If
                Else
                    '***** status monitor prevalente *****
                    'daniele luglio 2013 bolgiano: i codici di errore istantanei NVL e NVH finiscono nel codice di errore NVA sulle medie
                    Select Case Trim(Status_5_Secondi(Ora, nn, secondi))
                        
                        Case "ERR"
                            Stato_Monitor(0, 0) = Stato_Monitor(0, 0) + 1
                        
                        Case "TZR"
                            Stato_Monitor(1, 0) = Stato_Monitor(1, 0) + 1
                        
                        Case "TSP"
                            Stato_Monitor(2, 0) = Stato_Monitor(2, 0) + 1
                        
                        Case "MAN"
                            Stato_Monitor(3, 0) = Stato_Monitor(3, 0) + 1
                        
                        Case "OFF"
                            Stato_Monitor(4, 0) = Stato_Monitor(4, 0) + 1
                            
                        Case "NVA"
                            Stato_Monitor(5, 0) = Stato_Monitor(5, 0) + 1
                        
                        Case "NVL"
                            'Alby Dicembre 2015
                            Stato_Monitor(6, 0) = Stato_Monitor(6, 0) + 1
                            'Stato_Monitor(5, 0) = Stato_Monitor(5, 0) + 1
                        
                        Case "NVH"
                            Stato_Monitor(7, 0) = Stato_Monitor(7, 0) + 1
                            'Stato_Monitor(5, 0) = Stato_Monitor(5, 0) + 1
                            
                        Case "TAR"
                            Stato_Monitor(8, 0) = Stato_Monitor(8, 0) + 1
                        
                        
                    End Select
                End If
            End If
        End If
        ContaTutti_5_secondi(Ora, nn, 0, nmedie) = ContaTutti_5_secondi(Ora, nn, 0, nmedie) + ContaTuttiSecondiMediaOra(Ora, nn, secondi)
        
    Next secondi
    
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiAggiornoContatori " + Error(Err), 1)
    Resume fine
fine:
    
End Sub

Sub ElaboraSalvaDatiStatoImpiantoPrevalente()

    On Error GoTo GestErrore
        
    '***** stato impianto prevalente *****
    If IngressoIMPIANTO > -1 Then
        
        For Ora = 0 To OraFine
            For nmedie = 1 To numeromedie(Ora)
                If ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie) = 0 Then
                    statoimp(Ora, nmedie) = -9999
                Else
                    'daniele luglio 2013 bolgiano: la media è valida sul totale campioni orari 720!!
                    If (PercRegime(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie) >= 0.7) Then
                        statoimp(Ora, nmedie) = 30
                    Else
                        StatiImpianto(0, 0) = PercMinTec(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(0, 1) = 31

                        StatiImpianto(1, 0) = PercSpegnimento(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(1, 1) = 32

                        StatiImpianto(2, 0) = PercManutenzione(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(2, 1) = 33

                        StatiImpianto(3, 0) = PercFermo(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(3, 1) = 34

                        StatiImpianto(4, 0) = PercGuasto(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(4, 1) = 35

                        StatiImpianto(5, 0) = PercAnomalo(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(5, 1) = 36

                        'luca gennaio 2015 aggiungo gestione stato impianto 37
                        StatiImpianto(5, 0) = PercPolveri(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(5, 1) = 37

                        'Federica marzo 2018 aggiungo gestione stato impianto 38
                        StatiImpianto(6, 0) = PercAltro(Ora, nmedie) / ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie)
                        StatiImpianto(6, 1) = 38
                        
                        '*** lo stato impianto prevalente è l'elemento 0 della matrice ordinata decrescente
                        'luca gennaio 2015 aggiungendo uno stato impianto bisogna aggiornare anche i parametri passati alla routine QuickSort(da 5 a 6)
                        'Call QuickSort(StatiImpianto, 5, True)
                        Call QuickSort(StatiImpianto, 6, True)
                        statoimp(Ora, nmedie) = StatiImpianto(0, 1)
                    End If
                End If
            Next nmedie
        Next Ora
    End If
    
    Exit Sub

GestErrore:
    Call WindasLog("ElaboraSalvaDatiStatoImpiantoPrevalente " + Error(Err), 1)
    Resume fine
fine:

End Sub

Sub ElaboraSalvaDatiNormalizza()

    Dim ValoreTQ As Double
    Dim H2O As Double
    Dim O2 As Double
    Dim T As Double
    Dim P As Double
    Dim Status As String
    Dim iIdx As Integer
    
    On Error GoTo GestErrore

    '***** normalizzazione delle medie *****
    For nn = 0 To gnNroParametriStrumenti
        For Ora = 0 To OraFine
            For nmedie = 1 To numeromedie(Ora)
                
                '**** normalizzazione media ****
                H2O = ScegliDato(Ora, IngressoH2O, nmedie, Status)
                If Status <> "VAL" And Status <> "AUX" Then
                    H2O = -9999
                'luca aprile 2017 QAL2 su H2O
                Else
                    Dim tempH2O As Double
                    tempH2O = CalcolaQAL2(IngressoH2O, H2O)
                    H2O = tempH2O
                End If
                
                O2 = ScegliDato(Ora, IngressoO2, nmedie, Status)
                If Status <> "VAL" And Status <> "AUX" Then
                    O2 = -9999
                'luca aprile 2017 QAL2 su O2
                Else
                    Dim tempO2 As Double
                    tempO2 = CalcolaQAL2(IngressoO2, O2)
                    O2 = tempO2
                End If
                
                T = ScegliDato(Ora, IngressoTemp, nmedie, Status)
                If Status <> "VAL" And Status <> "AUX" Then T = -9999
                
                P = ScegliDato(Ora, IngressoPress, nmedie, Status)
                If Status <> "VAL" And Status <> "AUX" Then P = -9999
                
                ValoreTQ = ScegliDato(Ora, nn, nmedie, Status)
                
                MedieOra(Ora, nn, 1, nmedie) = ElaborazioniDiLegge(ValoreTQ, H2O, O2, T, P, nn, Status)
                StsMedieOra(Ora, nn, 1, nmedie) = Status
                
                'Alby Gennaio 2016 Versalis Priolo escluso per velocizzare
                'scarto quadratico valori grezzi / normalizzati / stimati
                For xx = 0 To 2
                    DoEvents
                    If ContaOraOK(Ora, nn, xx, nmedie) > 1 Then
                        SumDevStd(xx) = 0
                        For iIdx = 0 To MaxDati - 1 'Considero tutti i campioni dell'ora
                              If ValIstPerScarto(Ora, nn, iIdx, xx) <> -9999 Then
                                  SumDevStd(xx) = SumDevStd(xx) + (ValIstPerScarto(Ora, nn, iIdx, xx) - MedieOra(Ora, nn, xx, nmedie)) ^ 2
                              End If
                        Next
                        StdDev(Ora, nn, xx, nmedie) = Sqr(SumDevStd(xx) / (ContaOraOK(Ora, nn, xx, nmedie) - 1))
                        Call InserisciDecimali(StdDev(Ora, nn, xx, nmedie), 2)
                    Else
                        StdDev(Ora, nn, xx, nmedie) = -9999
                    End If
                Next xx
          Next nmedie
            
        Next Ora
        
    Next nn
    
    Exit Sub

GestErrore:
    Call WindasLog("ElaboraSalvaDatiNormalizza " + Error(Err), 1)
    Resume Next

End Sub

'Federica settembre 2017 - Calcolo H2O su medie Ossigeni
Sub ElaboraSalvaDatiCalcolaH2O(Optional ByVal FormulaAlternativa As Boolean = False)
    
    Dim MediaO2Umido As Double
    Dim MediaO2 As Double
    Dim FlagO2Umido As String
    Dim FlagO2 As String
    Dim MediaH2O As Double
    
    On Error GoTo GestErrore
    
    'Esco perchè non è necessario il calcolo
    If (IngressoO2Umido = -1) Or (IngressoO2 = -1) Then
        Exit Sub
    End If
    
    For Ora = 0 To OraFine
        For nmedie = 1 To numeromedie(Ora)
            MediaO2Umido = MedieOra(Ora, IngressoO2Umido, 1, nmedie)
            MediaO2 = MedieOra(Ora, IngressoO2, 1, nmedie)
            FlagO2Umido = StsMedieOra(Ora, IngressoO2Umido, 1, nmedie)
            FlagO2 = StsMedieOra(Ora, IngressoO2, 1, nmedie)
            
            If (MediaO2Umido <> -9999) And (MediaO2 <> -9999) And (InStr("VAL AUX", FlagO2Umido) > 0) And (InStr("VAL AUX", FlagO2) > 0) Then
                If MediaO2 <> 0 Then
                    'Federica dicembre 2017 - Aggiunta formula alternativa
                    If FormulaAlternativa Then
                        MediaH2O = 100 * (MediaO2 - MediaO2Umido) / MediaO2
                    Else
                        MediaH2O = 100 - (MediaO2Umido / MediaO2 * 100)
                    End If
                    If MediaH2O < 0 Then MediaH2O = 0
                Else
                    MediaH2O = 0
                End If
                
                MedieOra(Ora, IngressoH2O, 1, nmedie) = MediaH2O
                StsMedieOra(Ora, IngressoH2O, 1, nmedie) = "VAL"
            Else
                MedieOra(Ora, IngressoH2O, 1, nmedie) = -9999
                StsMedieOra(Ora, IngressoH2O, 1, nmedie) = "NCX"
            End If
        Next nmedie
    Next Ora
    
    Exit Sub
GestErrore:
    Call WindasLog("CalcolaH2O: " & Error(Err()), 1)

End Sub

Sub ElaboraSalvaDatiConcludo()

    On Error GoTo GestErrore
    
    '***** salvataggio dei dati nel DB *****
    For Ora = 0 To OraFine
        For nmedie = 1 To numeromedie(Ora)
            Form1.Label1.Caption = StrLabel & " Ora:" & Str(Ora) & " - Media:" & Str(nmedie)
            For nn = 0 To gnNroParametriStrumenti
                
                'luca luglio 2017
                Call OPC.ChiudiOPC
                
                DoEvents
                'luca 08/1/2016 salvo solo se non sono il client
                If Not Client Then
                    Call ElaboraSalvaDatiSQL(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
                End If
                
                'Alby Dicembre 2015 se è in automatico da BFwincc e se l'ora è la precedente del giorno corrente
                'luca marzo 2017
                If UCase(Tabella) <> "WDS_10MINCO" And UCase(Tabella) <> "WDS_AUTO" Then
                    'If InStr(Command, "auto") > 0 And (Format(ElabDate, "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") And Format(Ora, "00") = Format(DateAdd("h", -1, Now), "hh")) Then
                    If InStr(Command, "auto") > 0 And Ora = OraFine And nmedie = numeromedie(Ora) Then
                        'Elaboro dati e aggiorno WinCC
                        DatiPerWinCC = True
                        Call InizializzaWinCC
                        'luca marzo 2017
                        'Call ElaboraSalvaDatiMedieNF("48", nn, ElabDate)
                        If UCase(Tabella) = "WDS_ELAB" Then Call ElaboraSalvaDatiMedieNF("48", nn, Elabdate)
                        Call ElaboraAggiornaMedia(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
                    Else
                        DatiPerWinCC = False
                        'luca marzo 2017
                        'If Ora = 23 Then
                        If Ora = 23 And nmedie = numeromedie(Ora) Then
                            'se rielabora una giornata che non è il giorno corrente NON aggiorna WinCC
                            'e elabora medie 48h, giornaliere e mensili solo all'ultima ore nel ciclo la 23 (0-23)
                            'luca marzo 2017
                            'Call ElaboraSalvaDatiMedieNF("48", nn, ElabDate)
                            If UCase(Tabella) = "WDS_ELAB" Then Call ElaboraSalvaDatiMedieNF("48", nn, Elabdate)
                            Call ElaboraAggiornaMedia(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
                        End If
                    End If
                'luca aprile 2017
                ElseIf UCase(Tabella) = "WDS_10MINCO" Then
                    If InStr(Command, "auto") > 0 And Ora = OraFine And nmedie = numeromedie(Ora) Then
                        DatiPerWinCC = True
                        Call InizializzaWinCC
                        Call ElaboraAggiornaMedie10minuti(Ora, nn, nmedie)
                    End If
                End If
            Next nn
        Next nmedie
    Next Ora
    
    '***** se l'elaborazione selezionata è per i minuti non salvo il file .MEDIE per ARPA *****
    If TipoMedia < 3 Then
        '***** salvataggio file .MEDIE modello 4343 ARPA Lombardia *****
        Call ElaboraSalvaDatiConcludoADM(Elabdate, numeromedie())
    End If
    
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiConcludo " + Error(Err), 1)

End Sub

