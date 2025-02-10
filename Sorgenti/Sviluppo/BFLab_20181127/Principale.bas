Attribute VB_Name = "Principale"
''''''Alby Gennaio 2014
Option Explicit

'*** GESTIONE ***
Dim PrimaVolta As Boolean
Dim ResetVariabili As Boolean
Dim PrimaLetturaDI As Boolean
Dim SuddividiOperazioni As Integer
'----------------------------

Sub AcquisisceDati()
'§ Acquisizione dei valori analogici dalle Tag del PLC

    Dim ValoreAcquisito As Double
    Dim iIndice As Integer
    Dim FondoScalaScadaDaUsare As Double    'Federica febbraio 2018

    On Error GoTo Gesterrore

    For iIndice = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(iIndice)
            If .Acquisizione Then
                If (.FSE <> .ISE) Then
                    'Federica dicembre 2017
                    If .NroMorsetto >= 0 Then 'Per parametri calcolati il canale è -9999
                        '*** Azzero lo stato ***
                        Status(0, iIndice) = "VAL"
                        Status(1, iIndice) = "VAL"
                        
                        '*** Valore grezzo ***
                        #If versione = 3 Then
                            'Alby Marzo 2018
                            ValoreAcquisito = LeggiTag(CStr(NumeroLinea) & " AI" & CStr(.NroMorsetto))
                        #Else
                            ValoreAcquisito = LeggiTag(CStr(NumeroLinea) & "_DB80_AI_" & Format(.NroMorsetto, "00"))
                        #End If
                        
                        '*** Fondo scala per ingegnerizzazione ***
                        FondoScalaScadaDaUsare = IIf(.IndiceDigitale2CampoScala >= 0, .FSI2, .FSI)
                        
                        '*** Ingegnerizzazione ***
                        ValIst(0, iIndice) = .ISI + (ValoreAcquisito - .ISE) * (FondoScalaScadaDaUsare - .ISI) / (.FSE - .ISE)
                        ValIst(3, iIndice) = ValIst(0, iIndice)
                        
                        'Alby Agosto 2017
                        '*** Fattore di conversione UdM ***
                        If .FattoreConversione > 0 Then ValIst(0, iIndice) = ValIst(0, iIndice) * .FattoreConversione
                        
                        'luca luglio 2017 saturo al limite superiore istantaneo impostato con codice VAH
                        If ValIst(0, iIndice) > Val(Replace(.LimiteSuperiore, ",", ".")) Then
                            ValIst(0, iIndice) = Val(Replace(.LimiteSuperiore, ",", "."))
                            ValIst(1, iIndice) = Val(Replace(.LimiteSuperiore, ",", "."))
                            Status(0, iIndice) = "VAH"
                            Status(1, iIndice) = "VAH"
                        End If
                        
                        'luca luglio 2017 saturo al limite inferiore senza cambiare codice di validità
                        If ValIst(0, iIndice) < Val(Replace(.LimiteInferiore, ",", ".")) Then
                            ValIst(0, iIndice) = Val(Replace(.LimiteInferiore, ",", "."))
                            ValIst(1, iIndice) = Val(Replace(.LimiteInferiore, ",", "."))
                        End If
                        'luca aprile 2017 QAL2 su tal quale
                        If .QAL2suTQ And Valido(Status(0, iIndice)) Then
                            If .m <> 0 Then
                                ValIst(0, iIndice) = ValIst(0, iIndice) * .m + .q
                                ValIst(1, iIndice) = ValIst(0, iIndice)
                            End If
                        End If
                    Else
                        'Se non ho un canale di acquisizione
                        ValIst(0, iIndice) = -9999
                        ValIst(1, iIndice) = -9999
                        Status(0, iIndice) = "ERR"
                        Status(1, iIndice) = "ERR"
                    End If
                    
                Else
                    'Se non è impostata una scala strumentale
                    ValIst(0, iIndice) = -9999
                    ValIst(1, iIndice) = -9999
                    Status(0, iIndice) = "ERR"
                    Status(1, iIndice) = "ERR"
                End If
            
            Else
                'Se il parametro non è acquisito
                ValIst(0, iIndice) = -9999
                ValIst(1, iIndice) = -9999
                Status(0, iIndice) = "ERR"
                Status(1, iIndice) = "ERR"
            End If
        End With
    Next iIndice

    Exit Sub
    
Gesterrore:
    Call WindasLog("AcquisisceDati " + Error(Err), 1, OPC)
    'Resume Next

End Sub

Private Sub InvalidaMisure()

    Dim ii As Integer
    Dim iIndice As Integer

    On Error GoTo Gesterrore
                        
    For iIndice = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(iIndice)
            For ii = 0 To nroDigitali
                If Trim(CodiceParametro_DI(ii)) <> "" Then
                    If InStr(Trim(.Invalida), Trim(CodiceParametro_DI(ii))) > 0 Then
                        If Valore_DI(ii) = 1 Then
                            'luca 05/10/2016 aggiungo gestione vari codici di validità in base a digitali
                            If InStr(NomeParametro_DI(ii), "COMUNICAZIONE") > 0 Then
                                Status(0, iIndice) = "OFF"
                                Status(1, iIndice) = "OFF"
                            ElseIf InStr(NomeParametro_DI(ii), "MANUTENZIONE") > 0 Then
                                Status(0, iIndice) = "MAN"
                                Status(1, iIndice) = "MAN"
                            ElseIf InStr(NomeParametro_DI(ii), "ZERO") > 0 And Status(0, iIndice) <> "OFF" And Status(0, iIndice) <> "MAN" Then
                                Status(0, iIndice) = "TZR"
                                Status(1, iIndice) = "TZR"
                            ElseIf InStr(NomeParametro_DI(ii), "SPAN") > 0 And Status(0, iIndice) <> "OFF" And Status(0, iIndice) <> "MAN" Then
                                Status(0, iIndice) = "TSP"
                                Status(1, iIndice) = "TSP"
                            'luca luglio 2017
                            ElseIf InStr(NomeParametro_DI(ii), "CALIBRAZIONE") > 0 And Status(0, iIndice) <> "OFF" And Status(0, iIndice) <> "MAN" And Status(0, iIndice) <> "TZR" And Status(0, iIndice) <> "TSP" Then
                                Status(0, iIndice) = "TAR"
                                Status(1, iIndice) = "TAR"
                            ElseIf Status(0, iIndice) <> "OFF" And Status(0, iIndice) <> "MAN" And Status(0, iIndice) <> "TZR" And Status(0, iIndice) <> "TSP" Then
                                Status(0, iIndice) = "ERR"
                                Status(1, iIndice) = "ERR"
                            End If
                        End If
                    End If
                End If
            Next ii
        End With
    Next
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("InvalidaMisure: " & Error(Err()), 1, "OPC")
End Sub

'Federica ottobre 2017
Private Sub CalcolaPercentuale(ByVal indice As Integer)
'§ Calcolo del valore percentuale della misura normalizzata rispetto al fondo scala.
'  per i trend istantanei

    Dim Perc As Double

    On Error GoTo Gesterrore
    
    If ValIst(1, indice) = -9999 Then
        Perc = 0
    Else
        Perc = ValIst(1, indice) / ParametriStrumenti(indice).FSI * 100
        ValPerc(1, indice) = Perc
    End If
        
    Exit Sub
Gesterrore:
    Call WindasLog("CalcolaPercentuale: " & Error(Err()), 1, "OPC")

End Sub

Public Function Valido(ByVal strStato As String) As Boolean
'§ Verifica se lo status della misura rientra nei valori considerati come validi

    Dim Ret As Boolean
    
    On Error GoTo Gesterrore
    
    Ret = (InStr(strValid, strStato) > 0)

    Valido = Ret
    
    Exit Function
Gesterrore:
    Call WindasLog("Valido: " & Error(Err()), 1, "OPC")

End Function

Public Function TrasformaInDbl(ByVal valore As Variant) As Double

    On Error GoTo Gesterrore
    
    If TypeName(valore) = "String" And valore = "" Then
        TrasformaInDbl = 0
        Exit Function
    End If
    
    'TrasformaInDbl = CDbl(Replace(valore, ".", ","))
    TrasformaInDbl = Val(Replace(valore, ",", "."))
    
    Exit Function
Gesterrore:

    Call WindasLog("TrasformaInDbl: " & Error(Err()), 1, "OPC")
    TrasformaInDbl = 0

End Function

Sub ChiamaBFdata()
'§ Esecuzione di BFData se è il momento di calcolare le medie

    Dim Abilita10minutiCO As Boolean

    On Error GoTo Gesterrore
    
    Abilita10minutiCO = CBool(Generiche(i10Minuti).Par)
    
    'luca aprile 2017
    Select Case OreSemiore
        Case TIPO_MEDIE_ORARIE
            If second(Now) > 0 Then
                If minute(Now) = 0 Then
                    If EseguiMedie Then
                        EseguiMedie = False
                        Shell "C:\Windas\Windas03" & CStr(NumeroLinea) & "\" & CStr(NumeroLinea) & "_BFData.exe auto 2"
                    End If
                Else
                    EseguiMedie = True
                End If
            End If
        Case TIPO_MEDIE_SEMIORARIE
            If second(Now) > 0 Then
                If minute(Now) = 0 Or minute(Now) = 30 Then
                    If EseguiMedie Then
                        EseguiMedie = False
                        If Abilita10minutiCO Then
                            Shell "C:\Windas\Windas03" & CStr(NumeroLinea) & "\" & CStr(NumeroLinea) & "_BFData.exe auto 1 0"
                        Else
                            Shell "C:\Windas\Windas03" & CStr(NumeroLinea) & "\" & CStr(NumeroLinea) & "_BFData.exe auto 1"
                        End If
                    End If
                Else
                    EseguiMedie = True
                End If
            
                If Abilita10minutiCO Then
                    If minute(Now) Mod 10 = 0 And minute(Now) <> 0 And minute(Now) <> 30 Then
                        If EseguitMedie10MinutiCO Then
                            EseguitMedie10MinutiCO = False
                            Shell "C:\Windas\Windas03" & CStr(NumeroLinea) & "\" & CStr(NumeroLinea) & "_BFData.exe auto 0"
                        End If
                    Else
                        EseguitMedie10MinutiCO = True
                    End If
                End If
            End If
    End Select
    
Exit Sub
    
Gesterrore:
    Call WindasLog("ChiamaBFdata " + Error(Err), 1, OPC)

End Sub

Function CodParametro(Identificativo) As Integer
'§ Estrae l'indice del parametro a partire dall'ID Database
    
    Dim iIDParametro As Integer

    'Alby Dicembre 2015
    On Error GoTo Gesterrore
    
    'Federica ottobre 2017 - Se viene passato -1 non devo cercare
    If Identificativo = -1 Then
        CodParametro = Identificativo
        Exit Function
    End If
    
    CodParametro = -1
    For iIDParametro = 0 To gnNroParametriStrumenti
        If Identificativo = ParametriStrumenti(iIDParametro).idDatabase Then
            CodParametro = iIDParametro
            Exit Function
        End If
    Next iIDParametro
    Call WindasLog("Attenzione parametro NON trovato", 1, OPC)
    
    Exit Function
    
Gesterrore:
    Call WindasLog("CodParametro: " + Error(Err), 1, OPC)

End Function

Function DecodificaStatoImpianto(CodiceStatoImpianto) As String
'§ Per vaersione = 1

    Dim StatoImpiantoAttuale As String
    
    'Alby Dicembre 2015
    On Error GoTo Gesterrore
    
    Select Case CodiceStatoImpianto
    
        Case 30
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(1).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
    
        Case 31
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(2).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
    
        Case 32
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(3).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
             
        Case 33
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(4).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
        
        Case 34
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(5).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
            
        Case 35
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(6).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
            
        Case 36
            StatoImpiantoAttuale = LeggiTag(Trim(Generiche(7).Testo))
            DecodificaStatoImpianto = StatoImpiantoAttuale
    
        Case Else
            DecodificaStatoImpianto = "Stato Impianto non configurato"
            
    End Select
    
    Exit Function
    
Gesterrore:
    Call WindasLog("DecodificaStatoImpianto " + Error(Err), 1, OPC)

End Function

Sub ControlloWatchdogPLC()
'§ Verifica la comunicazione con il PLC tramite Tag Watchdog

    On Error GoTo Gesterrore
    
    'lettura watchdog da PLC
    Static OldValore As Double
    Dim valore As Double
    Dim adesso As Date
    Static contAnomalia As Integer
    Static lastCheck As Date
    Static RecuperoFatto As Boolean
    Dim IP_PLC As String    'Federica luglio 2017
    
    adesso = Now
    IP_PLC = Trim(Generiche(iIP_PLC).Testo)    'Federica luglio 2017
    If IP_PLC <> "" Then
        If PingTest(IP_PLC) Then
            #If versione = 3 Then
                valore = LeggiTag(CStr(NumeroLinea) & " _AI23")
            #Else
                valore = LeggiTag(CStr(NumeroLinea) & "_DB80_WD_FROMPLC")
            #End If
        Else
            Call WindasLog("ControlloWatchdogPLC: Nessuna comunicazione con il PLC", 1, OPC)
            valore = 0
        End If
    Else
        Call WindasLog("ControlloWatchdogPLC: IP PLC non presente", 1, OPC)
        valore = 0
    End If
    If OldValore = valore Then
        'Il valore è fermo
        If contAnomalia < 60 Then
            contAnomalia = contAnomalia + 1
        End If
    Else
        contAnomalia = 0
    End If
    OldValore = valore
    lastCheck = adesso
    If contAnomalia >= 60 Then
        manValoreDigitale(99, 0) = 1
        RecuperoFatto = False
    Else
        manValoreDigitale(99, 0) = 0
        
        If RecuperoFatto Then Call RecuperoDatiADAM5560ReadFile(RecuperoFatto)
        
        'Alby Agosto 2017
        If Not RecuperoFatto Then
            'Federica settembre 2017
            Call RecuperoDatiADAM5560Result("TODO")
            
            Call WindasLog("Rientro da anomalia recupero dati", 0, OPC)
            RecuperoFatto = True
            
            If Dir(PathBFImport) <> "" Then
                Shell PathBFImport
            Else
                Call WindasLog("ControlloWatchdogPLC: Manca programma BFImport.", 1, OPC)
            End If
        End If
    End If
 
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloWatchdogPLC " + Error(Err), 1, OPC)
 
End Sub

Sub InizializzaSistema()
'§ Settaggi iniziali

    On Error GoTo Gesterrore
        
    '***** Impostazione valori globali *****
    'Federica settembre 2017 - Costruzione dei percorsi per il recupero dei dati
    PathFileImportResult = App.Path & "\BFImportResult.txt"
    PathBFImport = App.Path & "\" & CStr(NumeroLinea) & "_BFimport.exe"
 
    PrimaVolta = True
    
    '***** Acquisizione parametri di connessione *****
    Call GetConnectionParam
    Call CheckDBConnection  'Federica gennaio 2018
    
    '***** Lettura configurazioni da BFDesk *****
    Call LeggiConfigurazione7
    
    'Setto la Tag che indica la lettura corretta della configurazione
    ScriviTag CStr(NumeroLinea) & "_LeggiConfig", 1
    
    '***** inizializzazione variabili solo all' avvio *****
    If Not ResetVariabili Then
        ResetVariabili = True
        Call InizializzaVariabili
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("InizializzaSistema " + Error(Err), 1, OPC)

End Sub

'luca 21/07/2016
Sub ScriviDescrizioneMisure()
'§ Per versione = 1

    Dim iParametro As Integer
    Dim DescrizioneLinguaAttuale As String
    
    On Error GoTo Gesterrore
    
    'ciclo per i parametri
    For iParametro = 0 To gnNroParametriStrumenti
        'mi leggo la variabile dizionario, la quale ha il testo della descrizione della misura in base alla lingua selezionata
        DescrizioneLinguaAttuale = LeggiTag(ParametriStrumenti(iParametro).NomeTagDizionario)
        'lo salvo sulla variabile standard che utilizzerò nel WinCC
        ScriviTag CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(iParametro).CodiceParametro, "000") & "_DESCRIPTION_FULL", DescrizioneLinguaAttuale
    Next iParametro
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ScriviDescrizioneMisure " + Error(Err), 1, OPC)

End Sub

'Federica novembre 2017 - Descrizione misure presa dal database
Sub ScriviDescrizioneMisureWinCC()
'§ Per versione = 2

    Dim iParametro As Integer
    Dim strInizioTag As String
    
    On Error GoTo Gesterrore
    
    'ciclo per i parametri
    For iParametro = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(iParametro)
            strInizioTag = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000")
            ScriviTag strInizioTag & "_DESCRIPTION_FULL", .DescrParametro + IIf(.UnitaMisura <> "---", " (" & .UnitaMisura & ")", "")
            ScriviTag strInizioTag & "_DESCRIPTION_NAME", .DescrParametro
            ScriviTag strInizioTag & "_DESCRIPTION_UDM", IIf(.UnitaMisura <> "---", .UnitaMisura, "")
        End With
    Next iParametro
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ScriviDescrizioneMisureWinCC: " + Error(Err), 1, OPC)

End Sub

Sub Acquisisce()
'§ Sub principale chiamata una volta ogni 5 secondi

    Dim Lingua As Integer
    Dim SecondiLetti As Integer
    Dim iParametro As Integer
    Dim OldLingua As Integer
    Static GiaAcquisito As Boolean
    Static OldMinuto As String
    Dim ElencoMisureDaSimulare As String
    Dim ElencoMisureQAL3 As String
    
    On Error GoTo Gesterrore
                                                                            
    ElencoMisureDaSimulare = Trim(Generiche(iMisureSimulate).Testo)
    ElencoMisureQAL3 = Trim(Generiche(iMisureQAL3).Testo)
                                                                            
    DoEvents
    
    If ConnessioneValida Then Call CheckDBConnection    'Federica gennaio 2018  - Test validità connessione al database
    
    If (Not PrimaVolta) And (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
        Call WindasLog("Inizializza Sistema", 0, OPC)
        Call InizializzaSistema
        
        'Alby Agosto 2017
        If InStr(UCase(App.EXEName), "BFIMPORT") > 0 Then
            Call RecuperoDatiADAM5560
            End
        End If
    End If
    
    'luca 21/07/2016 evento di cambio lingua -> devo cambiare le descrizioni delle misure nella pagina principale (1 italiano - 2 inglese)
    'luca giugno 2017 versione 1 - SiCEMS
    '                 versione 2 - Windas
    #If versione = 1 Then
        Lingua = LeggiTag("Lingua")
        If Lingua <> OldLingua Then
            Call ScriviDescrizioneMisure
        End If
        OldLingua = Lingua
    #Else
        'Federica dicembre 2017 - Compongo le Tag per la descrizione delle analogiche, da usare nella supervisione
        Call ScriviDescrizioneMisureWinCC
    #End If
    
    SecondiLetti = second(Now)
    sec_ndx = SecondiLetti \ 5
    SecondiLetti = SecondiLetti Mod 5
    Select Case SecondiLetti
        Case 0, 1, 2
            If Not GiaAcquisito Then
                '**** Acquisizione dati da strumenti ****
                Call AcquisisceDati
                
                '**** Acquisizione dati stimati ****
                Call AcquisisceMisureStimate(ElencoMisureStimate)
                
                '**** Simulazione misure per test ****
                Call SimulaMisure(ElencoMisureDaSimulare)
                
                GiaAcquisito = True
            End If
        Case 3, 4
            GiaAcquisito = False
    End Select
        
    For iParametro = 0 To gnNroParametriStrumenti
        ValIst(1, iParametro) = ValIst(0, iParametro)
        Status(1, iParametro) = Status(0, iParametro)
    Next iParametro
    
    '***** GESTIONE STATO IMPIANTO ****
    'Call StatoImpianto
    'Federica settembre 2018 - Enel.SI Agordo >>>>>
    Call EnelStatoImpianto
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** ELABORAZIONI SU DIGITALI *****
    Call ElaboraDigitali
    
    Call InvalidaMisure
    
    '**** Parametri calcolati ****
    'Inserire qui le chiamate
    Call CalcolaH2O
    Call ENELCalcolaPortataDaDeltaP
    
    'luca aprile 2017
    For iParametro = 0 To gnNroParametriStrumenti
        'Call ElaborazioniDiLegge(iParametro, ValIst(0, iParametro), ValIst(1, iParametro), Status(0, iParametro), Status(1, iParametro), _
                ValIst(0, IngressoTemp), ValIst(0, IngressoPress), ValIst(0, IngressoH2O), ValIst(0, IngressoO2), Status(0, IngressoTemp), _
                Status(0, IngressoPress), Status(0, IngressoH2O), Status(0, IngressoO2))
        Call ElaborazioniDiLegge(iParametro, ValIst(1, iParametro), Status(1, iParametro))
    Next iParametro
    
    If Not Client Then
        Call SalvaDAT
        Call SalvaDatiElementariDB 'Nicolò Agosto 2016 Predispongo salvataggio db elementare su db
    End If
    
    '**** BFDATA ****
    Call ChiamaBFdata
    
    '***** SME CLOUD *****
    'Federica febbraio 2018 - E' la Sub che gestisce la differenza tra ore e semiore
    Call GestioneSMECloud
    
    '***** DO SUPERO MEDIE *****
    Call GestioneSMECloudDOAllarmiMedie
    
    '***** LETTURA CONFIGURAZIONE *****
    If OldMinuto <> Format(Now, "nn") Then
        OldMinuto = Format(Now, "nn")
        
        'Federica gennaio 2018 - Test validità connessione al database
        'Se manca la connessione eseguo il test una volta al minuto altrimenti BFLab si pianta
        If Not ConnessioneValida Then Call CheckDBConnection
        
        If ConnessioneValida Then Call LeggiConfigurazione7 'Federica gennaio 2018 - Rileggo la configurazione solo se ho il DB connesso
        'Apertura e chiusura porta PLC in S7
'        If Refresh = True Then
'            If ConnectPLC(1) = False Then
'                Call InizializzaProtocollo(1)
'                DaLeggereRegistriPLC = True
'            End If
'        Else
'            Call DisconnectAll
'        End If
    End If
    
    '***** WATCHDOG *****
    'Federica luglio 2018
    Call ControlloWatchdogPLC
    'Federica luglio 2018
    Call ControlloWatchdogDriver
    
    '***** HOT BACKUP *****
    If AbilitaHotBackup Then
        Call ControlloHotBackup
    Else
        IsMaster = True
    End If
    
    'Call GestionePLC(1)
    'Federica ottobre 2017 - Parametri da generiche
    Call ControlloTarature(ElencoMisureQAL3)
    Call SetVariabiliWinCC
    
    '***** CAMBIAMENTI SU CONFIGURAZIONE *****
    Call ControlloConfigurazione
    
    'luca luglio 2017
    '***** SONORO ALLARMI *****
    If GestioneSonoroAbilitata Then Call GestioneAllarmiSonori
    
Exit Sub
    
Gesterrore:
    Call WindasLog("Acquisisce " + Error(Err), 1, OPC)
    Resume Next
    
End Sub

'luca maggio 2018
Sub ControlloWatchdogDriver()
'§ Verifica la comunicazione con il PLC tramite Tag Watchdog

    On Error GoTo Gesterrore

    Static OldValore As Double
    Dim valore As Double
    Dim adesso As Date
    Static contAnomalia As Integer
    Static lastCheck As Date
    
    adesso = Now
    valore = LeggiTag(CStr(NumeroLinea) & " WATCHDOG")
       
    If OldValore = valore Then
        If contAnomalia < 60 Then
            contAnomalia = contAnomalia + 1
        End If
    Else
        contAnomalia = 0
    End If
    
    OldValore = valore
    lastCheck = adesso
    If contAnomalia >= 60 Then
        manValoreDigitale(999, 9) = 1
    Else
        manValoreDigitale(999, 9) = 0
    End If
 
    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloWatchdogDriver: " + Error(Err), 1, OPC)
 
End Sub

Sub SimulaMisure(ByVal ElencoParametri As String)
'§ procedura di simulazione delle acquisizioni

    Dim iIdx As Integer
    Dim i As Integer
    Dim Parametri() As String

    On Error GoTo Gesterrore
    
    If ElencoParametri = "" Then Exit Sub
    
    Parametri = Split(ElencoParametri, ";")
    For iIdx = 0 To gnNroParametriStrumenti
        For i = 0 To UBound(Parametri)
            If CodParametro(CInt(Parametri(i))) = iIdx Then
                If iIdx = IngressoO2 Then
                    ValIst(0, iIdx) = Rnd * 8
                Else
                    ValIst(0, iIdx) = Rnd * 10
                End If
                ValIst(1, iIdx) = ValIst(0, iIdx)
            End If
        Next i
    Next iIdx

Exit Sub

Gesterrore:
    Call WindasLog("SimulaMisure " + Error(Err), 1, OPC)

End Sub

Sub InizializzaVariabili()
'§ Azzeramento delle variabili

    Dim iParametro As Integer
        
    On Error GoTo Gesterrore
    
    'Variabili per BFData
    EseguiMedie = False
    EseguitMedie10MinutiCO = False  'luca aprile 2017
    
    For iParametro = 0 To gnNroParametriStrumenti
        Status(0, iParametro) = "ERR"
        Status(1, iParametro) = "ERR"
    Next iParametro

    Exit Sub
    
Gesterrore:
    Call WindasLog("InizializzaVariabili " + Error(Err), 1, OPC)
    
End Sub

'luca aprile 2017
Function ProiezioneMedia(MediaInCorso, UltimoValore, StatusMediaInCorso, StatusUltimoValore) As Double
'§ Usata per il calcolo della media in costruzione

    Dim Minuti As Integer
    Dim MinutiMancanti As Integer
    Dim Proiezione As Double

    'Alby Dicembre 2015
    On Error GoTo Gesterrore
    
    'luca 11/10/2016 inserisco controllo se media in corso e ultimo valore sono validi e diversi da -9999
    If MediaInCorso = -9999 Or UltimoValore = -9999 Or StatusMediaInCorso <> "VAL" Or InStr("VAL VAH", StatusUltimoValore) = 0 Then
        ProiezioneMedia = -9999
        Exit Function
    End If
        
    'luca marzo 2017
'    Minuti = Val(Format(Now, "nn"))
'    MinutiMancanti = 60 - Minuti
'    Proiezione = ((MediaInCorso * Minuti) + (UltimoValore * MinutiMancanti)) / 60
'    ProiezioneMedia = Proiezione
    
    If OreSemiore = TIPO_MEDIE_ORARIE Then
        Minuti = Val(Format(Now, "nn"))
        MinutiMancanti = 60 - Minuti
        Proiezione = ((MediaInCorso * Minuti) + (UltimoValore * MinutiMancanti)) / 60
        ProiezioneMedia = Proiezione
    ElseIf OreSemiore = TIPO_MEDIE_SEMIORARIE Then
        Minuti = IIf(Val(Format(Now, "nn")) < 30, Val(Format(Now, "nn")), Val(Format(Now, "nn") - 30))
        MinutiMancanti = 30 - Minuti
        Proiezione = ((MediaInCorso * Minuti) + (UltimoValore * MinutiMancanti)) / 30
        ProiezioneMedia = Proiezione
    End If

    Exit Function

Gesterrore:
    Call WindasLog("ProiezioneMedia ", 1, OPC)

End Function

Sub ElaboraDigitali()
'§ Acquisizione dei valori analogici dalle Tag del PLC
    
    Dim iIndice As Integer
    Dim iIDBit As Integer
    Dim BitRiordinato As Integer
    Dim ValGroup As Double
    Dim nOldValore As Double
    Dim adesso As Date
    Dim Ora As String
    Dim Data As String
    Dim nroByteDigitali As Integer   'Federica ottobre 2017
    Dim TagValoreDigitale As Integer
    
    On Error GoTo Gesterrore
    
    nroByteDigitali = CInt(Generiche(iNrByte).Par)
    
    #If versione = 3 Then
        'luca maggio 2018
        For iIndice = 0 To nroByteDigitali
            If NroMorsetto_DI(iIndice, 0) <> -9999 Then
                TagValoreDigitale = Val(LeggiTag(Trim(CStr(NumeroLinea) & " DI" & CStr(NroMorsetto_DI(iIndice, 0)))))
                If TagValoreDigitale <> -9999 Then manValoreDigitale(NroMorsetto_DI(iIndice, 0), 0) = TagValoreDigitale
            End If
        Next iIndice
    
    #Else
        '***** lettura byte *****
        For iIndice = 0 To nroByteDigitali
            ValGroup = LeggiTag(CStr(NumeroLinea) & "_DB80_DI_" & Format(iIndice, "00"))
            #If versione = 1 Then
                For iIDBit = 0 To 15
                    'Alby Giugno 2014 ristrutturazione per i bit acquisiti da porta word da 16 bit
                    'Alby Ottobre 2014 nella versione Protec dovrebbero essere già giusti
                    BitRiordinato = iIDBit
                    'Alby Ottobre 2014 versione Robert
                    If iIDBit >= 8 Then
                        BitRiordinato = iIDBit - 8
                    Else
                        BitRiordinato = iIDBit + 8
                    End If
                    If (ValGroup And 2 ^ iIDBit) <> 0 Then
                        manValoreDigitale(iIndice, BitRiordinato) = 1
                    Else
                        manValoreDigitale(iIndice, BitRiordinato) = 0
                    End If
                Next iIDBit
            #Else
                'Federica luglio 2017 lettura di un solo byte
                For iIDBit = 0 To 7
                    If (ValGroup And 2 ^ iIDBit) <> 0 Then
                        manValoreDigitale(iIndice, iIDBit) = 1
                    Else
                        manValoreDigitale(iIndice, iIDBit) = 0
                    End If
                Next iIDBit
            #End If
        Next
    #End If

    If nroDigitali > -1 Then
        For iIndice = 0 To nroDigitali
            nOldValore = Valore_DI(iIndice)
            Valore_DI(iIndice) = manValoreDigitale(NroMorsetto_DI(iIndice, 0), NroMorsetto_DI(iIndice, 1))
            
            If Not StatoLogico_DI(iIndice) Then Valore_DI(iIndice) = Abs(Valore_DI(iIndice) - 1)
            
            'luca 08/11/2016 il client non deve salvare allarmi
            If Not Client Then
               If PrimaLetturaDI Then
                   If nOldValore <> Valore_DI(iIndice) Then
                       If (Priorita_DI(iIndice) And 1) <> 0 Then
                           adesso = Now
                           Ora = Format(adesso, "hh.mm.ss")
                           Data = Format(adesso, "yyyymmdd")
    
                           If Valore_DI(iIndice) <> 0 Then
                               'luca 20/07/2016 reimposto versione con rientro
                               Call AllarmiSalva(Ora, Data, gsClienteDi, Trim(CodiceParametro_DI(iIndice)), Trim(NomeParametro_DI(iIndice)), Trim(Testo1_DI(iIndice)), Trim(Famiglia_DI(iIndice)))
                               'luca luglio 2017 se abilitato il sonoro, all'uscita dell'allarme disabilito la tacitazione forzata
                               If Sonoro_DI(iIndice) = 1 Then ScriviTag "DisattivazioneAllarmiSonori", 0
                           Else
                               'Alby Dicembre 2015
                               Call AllarmiRientro(Trim(CodiceParametro_DI(iIndice)))
                           End If
                       End If
                   End If
               End If
            End If
        Next
        PrimaLetturaDI = True
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ElaboraDigitali " + Error(Err), 1, OPC)
    Resume Next

End Sub

Sub SetVariabiliWinCC()
'§ Scrittura delle Tag per la visualizzazione dei dati in HMI

    Dim iIndice As Integer
    Dim IndiceTipoDato As Integer   'luca 06/09/2016 gestisco ambedue le validità
    'Alby Gennaio 2014
    Dim inizio As Integer
    Dim fine As Integer
    Dim IP_PLC As String    'Federica luglio 2017
    Dim strInizioTag As String  'Federica febbraio 2018
    
    Static ValoreWD As Integer
    
    On Error GoTo Gesterrore
    
    IP_PLC = Trim(Generiche(iIP_PLC).Testo)    'Federica luglio 2017
    
    'luca 20/07/2016 uso nuove tag standard
    ScriviTag CStr(NumeroLinea) & ".DataOra_BFLab", Format(Now, "dd/mm/yyyy hh:nn:ss")
    'Federica gennaio 2018 - Scrivo il nome dell'impianto nella Tag per usarla nella supervisione
    ScriviTag CStr(NumeroLinea) & ".NOME_IMPIANTO", NomeLinea
    
    inizio = 0
    fine = gnNroParametriStrumenti
    
    'Alby Settembre 2014
    If (SuddividiOperazioni Mod 2) = 0 Then
        For iIndice = inizio To fine
            strInizioTag = CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(iIndice).CodiceParametro, "000")
            'luca 20/07/2016 uso nuove tag standard
            ScriviTag strInizioTag & "_IST", ValIst(0, iIndice)
            ScriviTag strInizioTag & "_ISTN", ValIst(1, iIndice)
            'luca 06/09/2016 scrittura validità grezza su tag WinCC
            ScriviTag strInizioTag & "_IST_VAL", Status(0, iIndice)
             'luca 06/09/2016 scrittura validità elaborata su tag WinCC
            ScriviTag strInizioTag & "_ISTN_VAL", Status(1, iIndice)
            'Federica ottobre 2017 scrittura tag per misura in percentuale
            ScriviTag strInizioTag & "_ISTN_PERC", ValPerc(1, iIndice)
            
            For IndiceTipoDato = 0 To 1
                Select Case IndiceTipoDato
                    Case 0
                        'luca luglio 2017
                        If Valido(Status(IndiceTipoDato, iIndice)) Then
                            'dato istantaneo grezzo valido -> tag VAL_VIS validità 0 default
                            ScriviTag strInizioTag & "_IST_VAL_VIS", 0
                            
                            'soglie attenzione - allarme dato istantaneo (lavora su soglia attenzione - allarme) (tag _VIS: attenzione -> giallo - allarme -> rosso)
                            ScriviTag strInizioTag & "_IST_VIS", DeterminaColore(ValIst(IndiceTipoDato, iIndice), iIndice, 0)
                        Else
                            'dato istantaneo grezzo non valido -> tag _VIS validità 2 (sfondo rosso, testo bianco)
                            ScriviTag strInizioTag & "_IST_VAL_VIS", 2
                        End If
                    Case 1
                        'luca luglio 2017
                        If Valido(Status(IndiceTipoDato, iIndice)) Then
                            'dato istantaneo grezzo valido -> tag VAL_VIS validità 0 default
                            ScriviTag strInizioTag & "_ISTN_VAL_VIS", 0
                            
                            'soglie attenzione - allarme dato istantaneo (lavora su soglia attenzione - allarme) (tag _VIS: attenzione -> giallo - allarme -> rosso)
                            ScriviTag strInizioTag & "_ISTN_VIS", DeterminaColore(ValIst(IndiceTipoDato, iIndice), iIndice, 0)
                        Else
                            'dato istantaneo grezzo non valido -> tag _VIS validità 2 (sfondo rosso, testo bianco)
                            ScriviTag strInizioTag & "_ISTN_VAL_VIS", 2
                        End If
                End Select
            Next IndiceTipoDato
        Next iIndice
    End If
        
    If SuddividiOperazioni = 1 Or SuddividiOperazioni = 10 Then
        Call SetVariabiliWinCCinCostruzione(inizio, fine)
    End If
    
    'Alby Gennaio 2014
    If ((SuddividiOperazioni Mod 2) = 1) And (SuddividiOperazioni <> 1) Then
        For iIndice = 0 To nroDigitali
            'luca 21/07/2016 uso tag standard
            ScriviTag CStr(NumeroLinea) & ".DM" & Format(CodiceParametro_DI(iIndice), "000") & "_IST", Valore_DI(iIndice)
        Next iIndice
    End If
    
    'Alby Gennaio 2014
    SuddividiOperazioni = SuddividiOperazioni + 1
    If SuddividiOperazioni > 20 Then SuddividiOperazioni = 0
    
    'luca 21/07/2016 uso tag standard e aggiungo scrittura codice numerico
    ScriviTag CStr(NumeroLinea) & ".STATO_IMPIANTO", ValIst(0, IngressoStatoImpianto)
    'Federica novembre 2017 - Versione 1 e 2 usano le stesse Generiche
    ScriviTag CStr(NumeroLinea) & ".STATO_IMPIANTO_STR", DecodificaStatoImpiantoWinCC()
    
    'luca 31/10/2016 watchdog to PLC
    'luca 08/11/2016 non scrivo se sono il client
    'Alby Maggio 2018 da verificare
    'luca maggio 2018
    If (Not Client) And IsMaster Then
        'Federica luglio 2017
        'Verifico la comunicazione con il PLC prima di scrivere
        If IP_PLC <> "" Then
            If PingTest(IP_PLC) Then
                If AbilitaWatchdogPLC Then ScriviTag CStr(NumeroLinea) & " AO0", Val(Format(Now, "ss"))
            Else
                Call WindasLog("SetVariabiliWinCC: Nessuna comunicazione con il PLC", 1, OPC)
            End If
        Else
            Call WindasLog("SetVariabiliWinCC: IP PLC non presente", 1, OPC)
        End If
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("SetVariabiliWinCC " + Error(Err), 1, OPC)

End Sub

'Federica novembre 2017 - Parametrizzazione delle Descrizioni stato impianto WinCC
Public Function DecodificaStatoImpiantoWinCC() As String

    Dim iIndice As Integer
    
    On Error GoTo Gesterrore
    
    For iIndice = 10 To 19
        If Val(Right(Trim(CStr(Generiche(iIndice).Descrizione)), 2)) = ValIst(0, IngressoStatoImpianto) Then
            DecodificaStatoImpiantoWinCC = Trim(Generiche(iIndice).Testo)
            Exit For
        End If
    Next iIndice
    
    If iIndice > 19 Then DecodificaStatoImpiantoWinCC = "Stato Impianto Non Definito"
    
    Exit Function
    
Gesterrore:
    Call WindasLog("DecodificaStatoImpiantoWinCC: " & Error(Err()), 1, "OPC")
    
End Function

Sub SetVariabiliWinCCinCostruzione(inizio, fine)
'§ Scrittura delle Tag per la visualizzazione dei dati in costruzione in HMI

    Dim iIndice As Integer
    Dim DatoPrevisionale As Double  'Alby Febbraio 2016
    'luca aprile 2017 2017
    Dim TagMediaInCorso As String
    Dim TagVisualizzazioneMediaInCorso As String
    Dim TagIDMediaInCorso As String
    Dim TagValiditaMediaInCorso As String
    Dim TagVisualizzazioneValiditaMediaInCorso As String
    Dim TagMediaPrevisionale As String
    Dim TagVisualizzazioneMediaPrevisionale As String
    Dim strInizioTag As String

    On Error GoTo Gesterrore

    For iIndice = inizio To fine
        strInizioTag = CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(iIndice).CodiceParametro, "000")
    
        'luca aprile 2017
        Select Case OreSemiore
            Case TIPO_MEDIE_ORARIE
                TagMediaInCorso = strInizioTag & "_MONC"
                TagVisualizzazioneMediaInCorso = strInizioTag & "_MONC_VIS"
                TagIDMediaInCorso = strInizioTag & "_MONC_ID"
                TagValiditaMediaInCorso = strInizioTag & "_MONC_VAL"
                TagVisualizzazioneValiditaMediaInCorso = strInizioTag & "_MONC_VAL_VIS"
                TagMediaPrevisionale = strInizioTag & "_MONP"
                TagVisualizzazioneMediaPrevisionale = strInizioTag & "_MONP_VIS"
            Case TIPO_MEDIE_SEMIORARIE
                TagMediaInCorso = strInizioTag & "_MSNC"
                TagVisualizzazioneMediaInCorso = strInizioTag & "_MSNC_VIS"
                TagIDMediaInCorso = strInizioTag & "_MSNC_ID"
                TagValiditaMediaInCorso = strInizioTag & "_MSNC_VAL"
                TagVisualizzazioneValiditaMediaInCorso = strInizioTag & "_MSNC_VAL_VIS"
                TagMediaPrevisionale = strInizioTag & "_MSNP"
                TagVisualizzazioneMediaPrevisionale = strInizioTag & "_MSNP_VIS"
        End Select
        
        'Media in corso
        ScriviTag TagMediaInCorso, MediaOraInCorso(1, iIndice)
        ScriviTag TagVisualizzazioneMediaInCorso, DeterminaColore(MediaOraInCorso(1, iIndice), iIndice, 0)
        ScriviTag TagIDMediaInCorso, ID_MediaOraInCorso(1, iIndice)
        ScriviTag TagValiditaMediaInCorso, StatusMediaOraInCorso(1, iIndice)
        'TODO: Anche VAH???
        If StatusMediaOraInCorso(1, iIndice) = "VAL" Then
            ScriviTag TagVisualizzazioneValiditaMediaInCorso, 0
        Else
            ScriviTag TagVisualizzazioneValiditaMediaInCorso, 2
        End If
        
        'Media previsionale
        DatoPrevisionale = ProiezioneMedia(MediaOraInCorso(1, iIndice), ValIst(1, iIndice), StatusMediaOraInCorso(1, iIndice), Status(1, iIndice))
        ScriviTag TagMediaPrevisionale, DatoPrevisionale
        ScriviTag TagVisualizzazioneMediaPrevisionale, DeterminaColore(DatoPrevisionale, iIndice, 0)
        
    Next iIndice

    Exit Sub
    
Gesterrore:
    Call WindasLog("SetVariabiliWinCCinCostruzione " + Error(Err), 1, OPC)

End Sub

Function DeterminaColore(valore, indice, Tipo) As Integer
'§ Determina il colore con cui viasualizzare i dati in HMI

    'Alby Febbraio 2016
    On Error GoTo Gesterrore
    
    'luca 06/09/2016 uso valori numerici (allineamento a WPF)
    DeterminaColore = 0
    If valore = -9999 Then
        'luca 06/09/2016
        DeterminaColore = 1
        Exit Function
    End If
    
    Select Case Tipo
        Case 0      'su limiti medi orari / istantanei ***************************************************
            If ParametriStrumenti(indice).SogliaAttenzione > 0 Then
                If valore > ParametriStrumenti(indice).SogliaAttenzione Then DeterminaColore = 5
            End If
            If ParametriStrumenti(indice).SogliaAllarme > 0 Then
                If valore > ParametriStrumenti(indice).SogliaAllarme Then DeterminaColore = 2
            End If
    
        Case 1      'su limiti medi giornalieri *********************************************
            If ParametriStrumenti(indice).SogliaAttenzioneGiornaliera > 0 Then
                If valore > ParametriStrumenti(indice).SogliaAttenzioneGiornaliera Then DeterminaColore = 5
            End If
            If ParametriStrumenti(indice).SogliaAllarmeGiornaliera Then
                If valore > ParametriStrumenti(indice).SogliaAllarmeGiornaliera Then DeterminaColore = 2
            End If
    End Select

    Exit Function

Gesterrore:
    Call WindasLog("DeterminaColore " + Error(Err), 1, OPC)

End Function

Sub StatoImpianto()
'§ Determinazione stato impianto da digitale

    Dim CodiceStatoImpianto As Integer
    Dim iIndice As Integer
    
    On Error GoTo Gesterrore
        
    If IngressoStatoImpianto = -1 Then Exit Sub
    
    'Federica novembre 2017 - Gestione digitali stato impianto da configurazione
    CodiceStatoImpianto = 34
    For iIndice = 10 To 19
        If (Generiche(iIndice).Par <> "---") And (Generiche(iIndice).Par <> vbEmpty) Then
            If Valore_DI(CInt(Generiche(iIndice).Par)) = 1 Then
                CodiceStatoImpianto = Val(Right(Trim(CStr(Generiche(iIndice).Descrizione)), 2))
                Exit For
            End If
        End If
    Next iIndice

    If Trim(Generiche(iMisureSimulate).Testo) <> "" Then
        'Se ho misure simulate forzo anche lo stato impianto
        ValIst(0, IngressoStatoImpianto) = 30
    Else
        ValIst(0, IngressoStatoImpianto) = CodiceStatoImpianto
    End If
    Status(0, IngressoStatoImpianto) = "VAL"
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("StatoImpianto " + Error(Err), 1, OPC)

End Sub

'luca aprile 2017 AGGIORNAMENTO tra calcolo media calcolata progressiva dagli elaborati (in costruzione) e metodo con il ricalcolo a fine ora i valori possono variare di molto (con QAL2 fa somma e sottrazioni)
'modifico in modo da passare i valori alla routine per richiamarla anche con la media in costruzione
'Sub ElaborazioniDiLegge(iIDParametro As Integer, ValoreTQ, ValoreNorm, StatusTQ, StatusNorm, ValoreTemp, ValorePress, ValoreH2O, ValoreO2, statusTemp, statusPress, statusH2O, statusO2)
Sub ElaborazioniDiLegge(iIDParametro As Integer, ValoreNorm, StatusNorm, Optional UsaMedie = False)
'§ Normalizzazione dei valori grezzi

    Dim tempH2O, tempO2, limO2, ValoreTemp, ValorePress, ValoreH2O, ValoreO2, ValoreTQ As Double
    Dim O2RIF As Integer
    Dim statusTemp, statusPress, statusH2O, statusO2 As String
    
    On Error GoTo Gesterrore
        
    'Leggo O2 di riferimento dalla configurazione
    O2RIF = CInt(Generiche(iO2RIF).Par)
    
    'Parto dal valore TQ
    ValoreTQ = IIf(UsaMedie, MediaOraInCorso(0, iIDParametro), ValIst(0, iIDParametro))
    
    With ParametriStrumenti(iIDParametro)
        If .Acquisizione And ValoreTQ <> -9999 Then 'Se è un parametro acquisito ed è stato acquisaito
            
            'Parto dal TQ
            ValoreNorm = IIf(UsaMedie, MediaOraInCorso(0, iIDParametro), ValIst(0, iIDParametro))
            StatusNorm = IIf(UsaMedie, StatusMediaOraInCorso(0, iIDParametro), Status(0, iIDParametro))
            
            '***** QAL2 *****
            'luca aprile 2017
            If Not .QAL2suTQ Then
                If .m <> 0 Then ValoreNorm = ValoreNorm * .m + .q
            End If
            
            '***** Normalizzazione in Temperatura e Pressione *****
            If InStr(.Elaborazioni, "N") <> 0 Then
                '***** Temperatura *****
                If IngressoTemp <> -1 Then
                    '***** Invalido la misura da correggere se la temperartura è invalida *****
                    'luca luglio 2017
                    If Not Valido(statusTemp) Then
                        StatusNorm = "NCT"
                    Else
                        ValoreTemp = IIf(UsaMedie, MediaOraInCorso(0, IngressoTemp), ValIst(0, IngressoTemp))
                        statusTemp = IIf(UsaMedie, StatusMediaOraInCorso(0, IngressoTemp), Status(0, IngressoTemp))
                        If ValoreTemp > 0 Then
                            If iIDParametro = IngressoPortata Then
                                ValoreNorm = ValoreNorm * (273.15 / (ValoreTemp + 273.15))
                            Else
                                ValoreNorm = ValoreNorm * ((ValoreTemp + 273.15) / 273.15)
                            End If
                        End If
                    End If
                End If
                
                '***** Pressione *****
                If IngressoPress <> -1 Then
                    
                    '***** Invalido la misura da correggere se la temperartura è invalida *****
                    'luca luglio 2017
                    If Not Valido(statusPress) Then
                        StatusNorm = "NCP"
                    Else
                        ValorePress = IIf(UsaMedie, MediaOraInCorso(0, IngressoPress), ValIst(0, IngressoPress))
                        statusPress = IIf(UsaMedie, StatusMediaOraInCorso(0, IngressoPress), Status(0, IngressoPress))
                        If ValorePress > 0 Then
                            If iIDParametro = IngressoPortata Then
                               ValoreNorm = ValoreNorm * (ValorePress / 1013.25)
                            Else
                                ValoreNorm = ValoreNorm * (1013.25 / ValorePress)
                            End If
                        End If
                    End If
                End If
            
            End If
        
            '***** Riporto al secco *****
            If InStr(.Elaborazioni, "S") <> 0 Then
                '***** Umidità *****
                If IngressoH2O <> -1 Then
                    '***** Invalido la misura da correggere se la temperartura è invalida *****
                    'luca luglio 2017
                    If Not Valido(statusH2O) Then
                        StatusNorm = "NCU"
                    Else
                        ValoreH2O = IIf(UsaMedie, MediaOraInCorso(0, IngressoH2O), ValIst(0, IngressoH2O))
                        statusH2O = IIf(UsaMedie, StatusMediaOraInCorso(0, IngressoH2O), Status(0, IngressoH2O))
                        
                        'luca aprile 2017 QAL2 su H2O
                        tempH2O = CalcolaQAL2(IngressoH2O, ValoreH2O)
                        'luca aprile 2017
                        If (tempH2O > 0) And (tempH2O < 100) Then
                            If iIDParametro = IngressoPortata Then
                                ValoreNorm = ValoreNorm * (100 - tempH2O) / 100
                            Else
                                ValoreNorm = ValoreNorm * 100 / (100 - tempH2O)
                            End If
                        End If
                    End If
                End If
            End If
        
            '***** Correzione al valore noto di Ossigeno *****
            If InStr(.Elaborazioni, "C") <> 0 Then
                '***** Ossigeno *****
                If IngressoO2 <> -1 Then
                    
                    'luca luglio 2017
                    If Not Valido(statusO2) Then
                        StatusNorm = "NCO"
                    Else
                        ValoreO2 = IIf(UsaMedie, MediaOraInCorso(0, IngressoO2), ValIst(0, IngressoO2))
                        statusO2 = IIf(UsaMedie, StatusMediaOraInCorso(0, IngressoO2), Status(0, IngressoO2))
                        
                        'luca aprile 2017
                        tempO2 = CalcolaQAL2(IngressoO2, ValoreO2)
                        If tempO2 >= 0 Then
                            'Alby Dicembre 2015
                            limO2 = 20
                            If tempO2 > limO2 Then tempO2 = limO2
                            If iIDParametro = IngressoPortata Then
                                ValoreNorm = ValoreNorm * (21 - tempO2) / (21 - O2RIF)
                            Else
                                ValoreNorm = ValoreNorm * (21 - O2RIF) / (21 - tempO2)
                            End If
                        End If
                    End If

                End If
            End If
            
            '***** sottrazione dell'intervallo di confidenza *****
            If .IntervalloConfidenza > 0 Then ValoreNorm = ValoreNorm - .IntervalloConfidenza
            
            '***** Limite di rilevabilita *****
            If .LimiteRilevabilita >= 0 Then
                If ValoreNorm < .LimiteRilevabilita Then ValoreNorm = .LimiteRilevabilita
            End If
        End If
    End With
                
    Exit Sub
    
Gesterrore:
    Call WindasLog("ElaborazionidiLegge errore: " + Error(Err), 1, OPC)
    'Resume Next
    
End Sub

'luca aprile 2017
Function CalcolaQAL2(indice As Integer, valore) As Double
'§ Calcolo QAL2 su TQ

    Dim temp As Double
    
    On Error GoTo Gesterrore
    
        temp = valore
        
        'se diverso da -9999 applica QAL2 altrimenti restituisce -9999
        If Not ParametriStrumenti(indice).QAL2suTQ Then
            If temp <> -9999 Then
                'QAL2
                If ParametriStrumenti(indice).m <> 0 Then
                    temp = temp * ParametriStrumenti(indice).m + ParametriStrumenti(indice).q
                End If
                
                'Intervallo di confidenza *****
                If ParametriStrumenti(indice).IntervalloConfidenza <> 0 Then
                    temp = temp - ParametriStrumenti(indice).IntervalloConfidenza
                End If
                
                '***** Limite di rilevabilita *****
                If ParametriStrumenti(indice).LimiteRilevabilita >= 0 Then
                    If temp < ParametriStrumenti(indice).LimiteRilevabilita Then
                        temp = ParametriStrumenti(indice).LimiteRilevabilita
                    End If
                End If
            End If
        End If
        
        CalcolaQAL2 = temp
        
    Exit Function

Gesterrore:
Call WindasLog("CalcolaQAL2: " + Error(Err), 1, OPC)

End Function

Function LeggiTag(NomeTag)
'§ Lettura dei valori dalle Tag

    On Error GoTo Gesterrore

    #If versione = 3 Then
        
        'Alby Marzo 2018
        Comunicator.CurrentItem = NomeTag
        LeggiTag = Comunicator.ItemValue
        
    #Else
        'Routine che gestisce le tag
        'sia in modalità WinCC script che VB6 OPC
            
        'Versione WinCC ActiveX
        LeggiTag = OPC.LeggiTagWinCC(NomeTag)
        
        'Script WinCC
        'HMIRuntime.Tags(NomeTag).Read

    #End If
    
    Exit Function
    
Gesterrore:
    Call WindasLog("LeggiTag " + Error(Err), 1, OPC)

End Function

Sub ScriviTag(NomeTag, Variabile)
'§ Scrittura dei valori delle Tag
    
    #If versione = 3 Then
            
        'Alby Marzo 2018
        On Error Resume Next
        
        Comunicator.AddItem NomeTag
        Comunicator.CurrentItem = NomeTag
        Comunicator.ItemValue = Variabile
    
    #Else
    
        'Routine che gestisce le tag
        'sia in modalità WinCC script che VB6 OPC
            
        'Versione WinCC ActiveX
        Call OPC.ScriviTagWinCC(NomeTag, Variabile)
        
        'Versione WinCC Script
        'HMIRuntime.Tags(NomeTag).Write Variabile
    #End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ScriviTag ", 1, OPC)
    
End Sub

Public Sub AllarmiSalva(ByVal Ora As String, ByVal Data As String, ByVal vsDirArchivio As String, ByVal Parameter As String, ByVal vsDescrizioneAllarme As String, ByVal StatoAllarme As String, Optional Tipo As Variant)
'§ Salvataggio degli allarmi nel DB quando si attivano

    Dim rsAllarmi As Object
    
    On Error GoTo GestErrSalvaAllarmi
        
    '***** Query di inserimento allarmi *****
    NewDataObj rsAllarmi
    With rsAllarmi
        strSQL = " INSERT INTO WDS_ALARM (AL_system,al_Station,al_Parameter,al_Hour,al_Date,AL_description,AL_group,AL_statusdesc) VALUES (" & _
             .ParSQLStr(Trim(gsImpianto)) & "," & _
             .ParSQLStr(Trim(gsClienteDi)) & "," & _
             .ParSQLStr(Trim(Parameter)) & "," & _
             .ParSQLStr(Trim(Ora)) & "," & _
             .ParSQLStr(Trim(Data)) & "," & _
             .ParSQLStr(Trim(vsDescrizioneAllarme)) & "," & _
             .ParSQLStr(Trim(Tipo)) & "," & _
             .ParSQLStr(Trim(StatoAllarme)) & ")"
       .ExecuteSql (strSQL)
    End With
    Set rsAllarmi = Nothing

    Exit Sub

GestErrSalvaAllarmi:
    Call WindasLog("Procedura AllarmiSalva " + Error(Err), 1, OPC)
End Sub

Sub AllarmiRientro(X As String)
'§ aggiornamento dell'allarme con il rientro

    Dim rsAllarmi As Object
    
    On Error GoTo Gesterrore
    
    '***** Inserimento nel record dell'ora di rientro allarme *****
    NewDataObj rsAllarmi
    With rsAllarmi
        strSQL = " UPDATE WDS_ALARM SET al_Hour2='" & Format(Now, "hh.nn.ss") & "',"
        strSQL = strSQL & "al_Date2='" & Format(Now, "yyyymmdd") & "' WHERE "
        strSQL = strSQL & "al_Station =" & .ParSQLStr(Trim(gsClienteDi)) & " AND "
        strSQL = strSQL & "al_Parameter=" & .ParSQLStr(Trim(X)) & " AND "
        strSQL = strSQL & "al_Date2 IS NULL AND "
        strSQL = strSQL & "al_Hour2 IS NULL"
       .ExecuteSql (strSQL)
    End With
    Set rsAllarmi = Nothing

    Exit Sub
    
Gesterrore:
    Call WindasLog("Errore in AllarmiRientro " + Error(Err), 1, OPC)

End Sub

Sub AcquisisceMisureStimate(ByVal ElencoParametri As String)
'§ Lettura del valore stimato inserito in configurazione
   
    Dim Parametri() As String
    Dim iIndice As Integer
    Dim CodiceParametro As Integer
   
    On Error GoTo Gesterrore
    
    'Alby Marzo 2018
    #If Not versione = 3 Then
    
        If ElencoParametri = "" Then Exit Sub
        
        Parametri = Split(ElencoParametri, ";")
        For iIndice = 0 To UBound(Parametri)
            CodiceParametro = CodParametro(CInt(Parametri(iIndice)))
            If CodiceParametro > 0 Then
                If ParametriStrumenti(CodiceParametro).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
                    ValIst(0, CodiceParametro) = TrasformaInDbl(ParametriStrumenti(CodiceParametro).OpzioniAcquisizione)
                    ValIst(1, CodiceParametro) = TrasformaInDbl(ParametriStrumenti(CodiceParametro).OpzioniAcquisizione)
                    Status(0, CodiceParametro) = "VAL"
                    Status(1, CodiceParametro) = "VAL"
                End If
            End If
        Next iIndice
    #End If
    
    Exit Sub

Gesterrore:
    Call WindasLog("AcquisisceMisureStimate " + Error(Err), 1, OPC)
    
End Sub

