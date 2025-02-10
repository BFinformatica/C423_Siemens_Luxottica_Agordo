Attribute VB_Name = "LeggiConfigurazione"
Option Explicit

Dim rsConfig As Object
    
Public Sub LeggiConfigurazione7()
    
    On Error GoTo Gesterrore
    
    '***** connessione a DB SQL *****
    NewDataObj rsConfig
    
    '***** Dati Linea *****
    'Federica gennaio 2018 - Lettura dati linea da DB
    strSQL = "select * from wds_gentab where gt_type='stations' and gt_code='" & gsClienteDi & "'"
    rsConfig.selectionfast strSQL
    If Not rsConfig.iseof Then
        NumeroLinea = rsConfig.getValue("gt_order")
        NomeLinea = rsConfig.getValue("gt_description")
    Else
        'Se manca il dato esco
        MsgBox "Manca numero linea per la stazione " & gsClienteDi & " nel DataBase. Esecuzione terminata!", vbCritical
        End
    End If
    
    'Setto la Tag che indica la lettura corretta della configurazione
    ScriviTag CStr(NumeroLinea) & "_LeggiConfig", 0
    
    Call LeggiConfigurazione7Misure
    Call LeggiConfigurazione7GenericheLinea
    Call LeggiConfigurazione7AssegnaParametri
    Call LeggiConfigurazione7Generiche
    Call LeggiConfigurazione7Digitali
    Call LeggiConfigurazione7AssegnaParametriWinCCSoglie
    Call LeggiConfigurazione7AssegnaParametriWinCCQAL2QAL3
    Call LeggiConfigurazione7AssegnaParametriWinCCValoreStimato
    Call LeggiConfigurazione7AssegnaParametriWinCCLimiti     'Federica luglio 2017
    Call LeggiConfigurazione7AssegnaParametriWinCCScale     'Federica ottobre 2017
    
    Set rsConfig = Nothing
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7 " + Error(Err), 1, OPC)
    Resume Next
    
End Sub

Public Sub LeggiConfigurazione7Misure()
    
    Dim tempArray() As String

    On Error GoTo Gesterrore
    
    '***** Azzeramento variabili *****
    gnNroParametriStrumenti = 0
    
    '***** Parametri generali *****
    strSQL = "SELECT was_measures.*, wds_gentab.gt_description " & _
             "FROM was_measures inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'params' AND cm_stationcode = '" & gsClienteDi & "' order by C1"
    rsConfig.selectionfast strSQL
    Do While Not rsConfig.iseof
        'Trasformo già i dati nel loro tipo
        ParametriStrumenti(gnNroParametriStrumenti).CodiceParametro = CStr(rsConfig.getValue("c1"))
        ParametriStrumenti(gnNroParametriStrumenti).NomeParametro = Trim(CStr(rsConfig.getValue("c2")))
        ParametriStrumenti(gnNroParametriStrumenti).DescrParametro = Trim(CStr(rsConfig.getValue("gt_description")))
        ParametriStrumenti(gnNroParametriStrumenti).UnitaMisura = Trim(CStr(rsConfig.getValue("c4")))
        ParametriStrumenti(gnNroParametriStrumenti).NroDecimali = CInt(rsConfig.getValue("c5"))
        ParametriStrumenti(gnNroParametriStrumenti).ISE = TrasformaInDbl(rsConfig.getValue("c6"))
        ParametriStrumenti(gnNroParametriStrumenti).FSE = TrasformaInDbl(rsConfig.getValue("c7"))
        ParametriStrumenti(gnNroParametriStrumenti).ISI = TrasformaInDbl(rsConfig.getValue("c8"))
        ParametriStrumenti(gnNroParametriStrumenti).FSI = TrasformaInDbl(rsConfig.getValue("c9"))
        ParametriStrumenti(gnNroParametriStrumenti).FSI2 = TrasformaInDbl(rsConfig.getValue("c10"))
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAttenzione = TrasformaInDbl(rsConfig.getValue("c11"))
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAllarme = TrasformaInDbl(rsConfig.getValue("c12"))
        ParametriStrumenti(gnNroParametriStrumenti).LimiteInferiore = TrasformaInDbl(rsConfig.getValue("c13"))
        ParametriStrumenti(gnNroParametriStrumenti).LimiteSuperiore = TrasformaInDbl(rsConfig.getValue("c14"))
        ParametriStrumenti(gnNroParametriStrumenti).LimiteInferioreOrario = TrasformaInDbl(rsConfig.getValue("c15"))
        ParametriStrumenti(gnNroParametriStrumenti).LimiteSuperioreOrario = TrasformaInDbl(rsConfig.getValue("c16"))
        ParametriStrumenti(gnNroParametriStrumenti).Acquisizione = CBool(rsConfig.getValue("c17"))
        ParametriStrumenti(gnNroParametriStrumenti).TipoAcquisizione = CInt(rsConfig.getValue("c18"))
        ParametriStrumenti(gnNroParametriStrumenti).OpzioniAcquisizione = rsConfig.getValue("c19")  'Se acquisito da strumento vale "Advantec", altrimenti trasformato in Double
        ParametriStrumenti(gnNroParametriStrumenti).TipoStrumento = rsConfig.getValue("c20")    'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).NroMorsetto = CInt(rsConfig.getValue("c21"))
        ParametriStrumenti(gnNroParametriStrumenti).SogliaValidazione = rsConfig.getValue("c22")    'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).idDatabase = CInt(rsConfig.getValue("c23")) 'Alby Dicembre 2015
        ParametriStrumenti(gnNroParametriStrumenti).Elaborazioni = CStr(rsConfig.getValue("c24"))
        'Alby Febbraio 2016
        ParametriStrumenti(gnNroParametriStrumenti).MaxIncremento = rsConfig.getValue("c25")    'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).MinEscursione = rsConfig.getValue("c26")    'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).MaxEscursione = rsConfig.getValue("c27")    'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaSemiorariaA = TrasformaInDbl(rsConfig.getValue("c28"))  'luca aprile 2017
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaOraria = TrasformaInDbl(rsConfig.getValue("c30"))
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaGiornaliera = TrasformaInDbl(rsConfig.getValue("c31"))
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMedia48H = rsConfig.getValue("c32")  'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaMensile = TrasformaInDbl(rsConfig.getValue("C33"))
        ParametriStrumenti(gnNroParametriStrumenti).Invalida = CStr(rsConfig.getValue("c36"))
        'Alby Febbraio 2016
        ParametriStrumenti(gnNroParametriStrumenti).LimiteRilevabilita = TrasformaInDbl(rsConfig.getValue("c40"))
        ParametriStrumenti(gnNroParametriStrumenti).m = TrasformaInDbl(rsConfig.getValue("c41"))
        ParametriStrumenti(gnNroParametriStrumenti).q = TrasformaInDbl(rsConfig.getValue("c42"))
        ParametriStrumenti(gnNroParametriStrumenti).Range = TrasformaInDbl(rsConfig.getValue("c43"))    'luca 25/07/2016 range di validià QAL2 e data QAL2
        ParametriStrumenti(gnNroParametriStrumenti).IntervalloConfidenza = TrasformaInDbl(rsConfig.getValue("c44"))
        ParametriStrumenti(gnNroParametriStrumenti).DataQAL2 = CStr(rsConfig.getValue("c45"))
        ParametriStrumenti(gnNroParametriStrumenti).ZeroSams = TrasformaInDbl(rsConfig.getValue("c46"))
        ParametriStrumenti(gnNroParametriStrumenti).SpanSams = TrasformaInDbl(rsConfig.getValue("c47"))
        ParametriStrumenti(gnNroParametriStrumenti).NomeTagDizionario = CStr(rsConfig.getValue("c51"))  'luca 21/07/2016
        ParametriStrumenti(gnNroParametriStrumenti).Precisione = rsConfig.getValue("c53")  'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).SogliaIstMin = rsConfig.getValue("c55") 'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).SogliaIstMax = rsConfig.getValue("c56") 'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaAnnuale = TrasformaInDbl(rsConfig.getValue("c57"))
        ParametriStrumenti(gnNroParametriStrumenti).ZeroTeorico = TrasformaInDbl(rsConfig.getValue("c58"))
        ParametriStrumenti(gnNroParametriStrumenti).SpanTeorico = TrasformaInDbl(rsConfig.getValue("c59"))
        'luca 05/10/2016 per gestione configurazione
        tempArray = Split(rsConfig.getValue("C60"), ";") '0 soglie, 1 QAL2/QAL3, 2 Valore Stimato
        ParametriStrumenti(gnNroParametriStrumenti).AttivaControlloConfigurazioneSoglie = CBool(tempArray(0))
        ParametriStrumenti(gnNroParametriStrumenti).AttivaControlloConfigurazioneQAL2QAL3 = CBool(tempArray(1))
        ParametriStrumenti(gnNroParametriStrumenti).AttivaControlloConfigurazioneValoreStimato = CBool(tempArray(2))
        
        ParametriStrumenti(gnNroParametriStrumenti).QAL2suTQ = CBool(IIf(rsConfig.getValue("c61") = "", 0, rsConfig.getValue("c61")))  'luca aprile 2017
        ParametriStrumenti(gnNroParametriStrumenti).ErroreZero = rsConfig.getValue("c62")   'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).ErroreSpan = rsConfig.getValue("c63")   'Non viene usato
        ParametriStrumenti(gnNroParametriStrumenti).CodiceMonitorIst_TQ = CStr(rsConfig.getValue("c65"))
        ParametriStrumenti(gnNroParametriStrumenti).CodiceMonitorMed_TQ = CStr(rsConfig.getValue("c66"))
        ParametriStrumenti(gnNroParametriStrumenti).CodiceMonitorMed_EL = CStr(rsConfig.getValue("c67"))
        ParametriStrumenti(gnNroParametriStrumenti).OrdineParametriADIADM = CInt(rsConfig.getValue("c68"))
            
        'Alby Marzo 2018
        #If Not versione = 3 Then
            ParametriStrumenti(gnNroParametriStrumenti).IndiceDigitale2CampoScala = CInt(rsConfig.getValue("C73"))  'Federica febbraio 2018
        #End If
        
        'luca 07/10/2015 soglia attenzione e allarme giornaliera
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAttenzioneGiornaliera = TrasformaInDbl(rsConfig.getValue("c75"))
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAllarmeGiornaliera = TrasformaInDbl(rsConfig.getValue("c76"))
        ParametriStrumenti(gnNroParametriStrumenti).LimConcMediaTrimestrale = TrasformaInDbl(rsConfig.getValue("c78"))
        ParametriStrumenti(gnNroParametriStrumenti).FattoreConversione = TrasformaInDbl(rsConfig.getValue("c80"))   'Alby Agosto 2017
        'Federica ottobre 2017
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAttenzioneMensile = TrasformaInDbl(rsConfig.getValue("L10"))
        ParametriStrumenti(gnNroParametriStrumenti).SogliaAllarmeMensile = TrasformaInDbl(rsConfig.getValue("L11"))
        
        'Federica giugno 2018
        ParametriStrumenti(gnNroParametriStrumenti).IndiceTagCollegataQAL2 = Val(rsConfig.getValue("C77"))

        gnNroParametriStrumenti = gnNroParametriStrumenti + 1
        rsConfig.MoveNext
    Loop
    
    'Alby Gennaio 2014
    gnNroParametriStrumenti = gnNroParametriStrumenti - 1
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7Misure: " + Error(Err), 1, OPC)
    Resume Next
    
End Sub

'luca 21/09/2016
Private Sub LeggiConfigurazione7AssegnaParametriWinCCSoglie()

    Dim i As Integer
    Dim strInizioTag As String
    
    On Error GoTo Gesterrore
    
    For i = 0 To gnNroParametriStrumenti
        'Federica dicembre 2017 - Rileggo solo se è stata aggoirnata altrimenti crea problemi nella supervisione
        'perchè riscrive il valore precedente mentre l'operatore sta inserendo i valori
        'If (ParametriStrumenti(i).AttivaControlloConfigurazioneSoglie) And (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
        'Federica ottobre 2018 - Disabilito la gestione della configurazione perchè salva la supervisione direttamente nel database
        If (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
            strInizioTag = "CONFIG" & CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(i).CodiceParametro, "000")
            ScriviTag strInizioTag & "_SATT", ParametriStrumenti(i).SogliaAttenzione
            ScriviTag strInizioTag & "_SALL", ParametriStrumenti(i).SogliaAllarme
            ScriviTag strInizioTag & "_SATT_GIORNO", ParametriStrumenti(i).SogliaAttenzioneGiornaliera
            ScriviTag strInizioTag & "_SALL_GIORNO", ParametriStrumenti(i).SogliaAllarmeGiornaliera
            ScriviTag strInizioTag & "_SATT_MESE", ParametriStrumenti(i).SogliaAttenzioneMensile
            ScriviTag strInizioTag & "_SALL_MESE", ParametriStrumenti(i).SogliaAllarmeMensile
        End If
    Next i
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametriWinCCSoglie " + Error(Err), 1, OPC)

End Sub

Private Sub LeggiConfigurazione7Generiche()
    
    On Error GoTo Gesterrore
    
    '**** Codice Impianto ****
    rsConfig.selectionfast "select gt_code from WDS_GENTAB where GT_TYPE = 'SYSTEMS'"
    gsImpianto = rsConfig.getValue("gt_code")
    
    '**** Directory di lavoro ****
    rsConfig.selectionfast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='DIR_LAV'"
    gsDirLavoro = rsConfig.getValue("gt_value")
    
    '**** Cartella per file DAT ****
    'luca 20/07/2016 ogni linea ha la sua cartella DATM dove salvare i DATM
    rsConfig.selectionfast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='DAT_FLD'"
    PathDAT = gsDirLavoro & "Windas03" & CStr(NumeroLinea) & "\" & rsConfig.getValue("gt_value")
    'Nicolò Agosto 2016
    PathDBElementare = gsDirLavoro & "Windas03" & CStr(NumeroLinea) & "\DB"
    
    'Alby Ottobre 2013 la tabella was_config viene allineata dal BFdatasync
    'spostato il parametro in wds_gentab
    rsConfig.selectionfast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='FUNCTION'"
    PCFunction = rsConfig.getValue("gt_value")
    
    'Alby Agosto 2017
    rsConfig.selectionfast "SELECT GT_VALUE FROM WDS_GENTAB WHERE GT_CODE='ADAMPATH'"
    AdamPath = rsConfig.getValue("GT_VALUE")
    
    'Federica Ottobre 2017
    rsConfig.selectionfast "SELECT GT_Code FROM wds_GenTab WHERE GT_Type = 'VALID' AND (GT_NumInt <> 0 AND GT_NumInt IS NOT NULL) "
    Do While Not rsConfig.iseof
      strValid = strValid & rsConfig.ParSQLStr(rsConfig.getValue("GT_Code")) & " "
      rsConfig.MoveNext
    Loop
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7Generiche " + Error(Err), 1, OPC)

End Sub

Private Sub LeggiConfigurazione7GenericheLinea()

    Dim indice As Integer

    On Error GoTo Gesterrore
    
    '***** generiche per linea *****
    'luca 20/07/2016 aggiungo indice
    rsConfig.selectionfast "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & gsClienteDi & "' ORDER BY CC_CODE"
    While (Not (rsConfig.iseof))
        indice = rsConfig.getValue("cc_code")
        Generiche(indice).Par = rsConfig.getValue("cc_value")
        Generiche(indice).Testo = Trim(rsConfig.getValue("cc_text"))
        Generiche(indice).Descrizione = Trim(rsConfig.getValue("cc_description"))
        
        rsConfig.MoveNext
    Wend
    
    'luca aprile 2017
    OreSemiore = CInt(Generiche(iOreSemiore).Par)
    'Federica ottobre 2017
    ElencoMisureStimate = Trim(Generiche(iMisureStimate).Testo)
    Nome_Impianto_4343 = Trim(Generiche(i4343_Impianto).Testo)
    Nome_Software_4343 = Trim(Generiche(i4343_SW).Testo)
    Nome_File_4343 = Trim(Generiche(i4343_File).Testo)
    AbilitaWatchdogPLC = CBool(Generiche(iWDPLC).Par)
    AbilitaHotBackup = CBool(Generiche(iHotBackup).Par)
    
    Exit Sub
Gesterrore:

    Call WindasLog("LeggiConfigurazione7GenericheLinea: " & Error(Err()), 1, "OPC")

End Sub

Private Sub LeggiConfigurazione7Digitali()

    Dim iIndice As Integer
    Dim CanaleDigitale() As String

    On Error GoTo Gesterrore
    
    'Azzeramento variabili
    nroDigitali = 0

    '***** configurazione digitali *****
    iIndice = 0
    strSQL = "SELECT was_digital.*, wds_gentab.gt_description, wds_gentab.gt_str5 " & _
             "FROM was_digital inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'alarm' AND cd_stationcode = '" & gsClienteDi & "' order by C1"
    rsConfig.selectionfast strSQL
    Do While Not rsConfig.iseof
        CodiceParametro_DI(iIndice) = rsConfig.getValue("c1")
        NomeParametro_DI(iIndice) = rsConfig.getValue("gt_description")
        Famiglia_DI(iIndice) = rsConfig.getValue("gt_str5")
        
        'Alby Ottobre 2014
        CanaleDigitale = Split(rsConfig.getValue("c3"), ".")
        NroMorsetto_DI(iIndice, 0) = CanaleDigitale(0)
        NroMorsetto_DI(iIndice, 1) = CanaleDigitale(1)
        
        Contatti_DI(iIndice) = rsConfig.getValue("c4")
        If rsConfig.getValue("c5") = "1" Then
            StatoLogico_DI(iIndice) = True
        Else
            StatoLogico_DI(iIndice) = False
        End If
        Colore0_DI(iIndice) = rsConfig.getValue("c6")
        Testo0_DI(iIndice) = rsConfig.getValue("c7")
        Colore1_DI(iIndice) = rsConfig.getValue("c8")
        Testo1_DI(iIndice) = rsConfig.getValue("c9")
        If rsConfig.getValue("c10") = "0" Then
            Priorita_DI(iIndice) = 0
        Else
            Priorita_DI(iIndice) = 1
        End If
        If rsConfig.getValue("c11") <> "0" Then
            Priorita_DI(iIndice) = Priorita_DI(iIndice) + 2
        End If
        
        Sonoro_DI(iIndice) = Val(rsConfig.getValue("c16"))  'luca luglio 2017
        IndiceDO(CanaleDigitale(0), CanaleDigitale(1)) = rsConfig.getValue("C17")   'Federica dicembre 2017
        
        iIndice = iIndice + 1
        rsConfig.MoveNext
    Loop
    nroDigitali = iIndice - 1
    
    '***** VERIFICO SE CI SONO ALLARMI CHE HANNO IL SONORO ABILITATO *****
    strSQL = "SELECT COUNT(*) AS Tot " & _
             "FROM was_digital " & _
             "WHERE cd_stationcode = '" & gsClienteDi & "' AND C16 = '1'"
    rsConfig.selectionfast strSQL
    If Not rsConfig.iseof Then GestioneSonoroAbilitata = (rsConfig.getValue("Tot") > 0)
    
    '***** VERIFICO SE CI SONO ALLARMI CHE HANNO LO SME CLOUD ABILITATO *****
    strSQL = "SELECT COUNT(*) AS Tot " & _
             "FROM was_digital " & _
             "WHERE cd_stationcode = '" & gsClienteDi & "' AND C14 = '1'"
    rsConfig.selectionfast strSQL
    If Not rsConfig.iseof Then GestioneSMECloudAbilitata = (rsConfig.getValue("Tot") > 0)
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7Digitali " + Error(Err), 1, OPC)
    Resume Next

End Sub

Private Sub LeggiConfigurazione7AssegnaParametri()

    On Error GoTo Gesterrore
        
    'Ingressi da configurare sempre in configurazione generiche di linea
    'Se non presenti mettere -1
    IngressoStatoImpianto = IIf(CInt(Generiche(iStatoImpianto).Par) = -1, -1, CodParametro(CInt(Generiche(iStatoImpianto).Par)))
    IngressoPortata = IIf(CInt(Generiche(iPortata).Par) = -1, -1, CodParametro(CInt(Generiche(iPortata).Par)))
    IngressoTemp = IIf(CInt(Generiche(iTemperatura).Par) = -1, -1, CodParametro(CInt(Generiche(iTemperatura).Par)))
    IngressoPress = IIf(CInt(Generiche(iPressione).Par) = -1, -1, CodParametro(CInt(Generiche(iPressione).Par)))
    IngressoH2O = IIf(CInt(Generiche(iH2O).Par) = -1, -1, CodParametro(CInt(Generiche(iH2O).Par)))
    IngressoO2 = IIf(CInt(Generiche(iO2).Par) = -1, -1, CodParametro(CInt(Generiche(iO2).Par)))
    IngressoO2Umido = IIf(CInt(Generiche(iO2Umido).Par) = -1, -1, CodParametro(CInt(Generiche(iO2Umido).Par)))
    IngressoNO = IIf(CInt(Generiche(iNO).Par) = -1, -1, CodParametro(CInt(Generiche(iNO).Par)))
    IngressoNOX = IIf(CInt(Generiche(iNOX).Par) = -1, -1, CodParametro(CInt(Generiche(iNOX).Par)))
    IngressoVelocita = IIf(CInt(Generiche(iVelocita).Par) = -1, -1, CodParametro(CInt(Generiche(iVelocita).Par)))
    IngressoNO2 = IIf(CInt(Generiche(iNO2).Par) = -1, -1, CodParametro(CInt(Generiche(iNO2).Par)))  'Federica gennaio 2018
    IngressoDeltaP = IIf(CInt(Generiche(iDeltaP).Par) = -1, -1, CodParametro(CInt(Generiche(iDeltaP).Par)))  'Federica gennaio 2018
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametri " + Error(Err), 1, OPC)

End Sub

'luca 25/07/2016
Private Sub LeggiConfigurazione7AssegnaParametriWinCCQAL2QAL3()

    Dim i As Integer
    Dim strInizioTag As String
    
    On Error GoTo Gesterrore
    
    For i = 0 To gnNroParametriStrumenti
        'Federica ottobre 2018 - Disabilito perchè il salavataggio viene fatto dalla supervisione direttamente nel database
        'If ParametriStrumenti(i).AttivaControlloConfigurazioneQAL2QAL3 Then
        If (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
            strInizioTag = CStr(NumeroLinea) & ".AM" & Format(i, "000")
            ScriviTag strInizioTag & "_QAL2_M", ParametriStrumenti(i).m
            ScriviTag strInizioTag & "_QAL2_Q", ParametriStrumenti(i).q
            ScriviTag strInizioTag & "_QAL2_RANGE", ParametriStrumenti(i).Range
            ScriviTag strInizioTag & "_QAL2_IC", ParametriStrumenti(i).IntervalloConfidenza
            If Len(ParametriStrumenti(i).DataQAL2) = 8 Then
                ScriviTag strInizioTag & "_QAL2_DATE", Right(ParametriStrumenti(i).DataQAL2, 2) & "/" & Mid(ParametriStrumenti(i).DataQAL2, 5, 2) & "/" & Left(ParametriStrumenti(i).DataQAL2, 4)
            Else
                'luca 25/07/2016 scrivo now perchè opc va in errore con stringa vuota
                ScriviTag strInizioTag & "_QAL2_DATE", Now
            End If
            ScriviTag strInizioTag & "_QAL3_ZEROREF", ParametriStrumenti(i).ZeroTeorico
            ScriviTag strInizioTag & "_QAL3_SPANREF", ParametriStrumenti(i).SpanTeorico
        End If
    Next i
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametriWinCCQAL2QAL3 " + Error(Err), 1, OPC)

End Sub

'luca 05/10/2016
Private Sub LeggiConfigurazione7AssegnaParametriWinCCValoreStimato()

    Dim i As Integer
    
    On Error GoTo Gesterrore
    
    For i = 0 To gnNroParametriStrumenti
        'Federica ottobre 2018 - Disabilito perchè il salavataggio viene fatto dalla supervisione direttamente nel database
        'If ParametriStrumenti(i).AttivaControlloConfigurazioneValoreStimato And (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
        If (LeggiTag(CStr(NumeroLinea) & "_LeggiConfig") = 0) Then
            If ParametriStrumenti(i).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
                 ScriviTag "CONFIG" & CStr(NumeroLinea) & ".AM" & Format(i, "000") & "_VALSTIMATO", TrasformaInDbl(ParametriStrumenti(i).OpzioniAcquisizione)
            End If
        End If
    Next i
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametriWinCCValoreStimato " + Error(Err), 1, OPC)

End Sub

'Federica luglio 2017
Private Sub LeggiConfigurazione7AssegnaParametriWinCCLimiti()

    Dim i As Integer
    Dim strInizioTag
    
    On Error GoTo Gesterrore
    
    For i = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(i)
            strInizioTag = CStr(NumeroLinea) & ".AM" & Format(i, "000")
            'Impostazione delle Tag per i limiti
            Call ScriviTag(strInizioTag & IIf(OreSemiore = TIPO_MEDIE_ORARIE, "_MOL", "_MSL"), IIf(OreSemiore = TIPO_MEDIE_ORARIE, .LimConcMediaOraria, .LimConcMediaSemiorariaA))
            Call ScriviTag(strInizioTag & IIf(OreSemiore = TIPO_MEDIE_ORARIE, "_MGOL", "_MGSL"), .LimConcMediaGiornaliera)
            Call ScriviTag(strInizioTag & IIf(OreSemiore = TIPO_MEDIE_ORARIE, "_MMOL", "_MMSL"), .LimConcMediaMensile)
        End With
    Next i

    Exit Sub
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametriWinCCLimiti: " & Error(Err()), 1, OPC)

End Sub

'Federica ottobre 2017
Private Sub LeggiConfigurazione7AssegnaParametriWinCCScale()

    Dim i As Integer
    
    On Error GoTo Gesterrore
    
    For i = 0 To gnNroParametriStrumenti
        'Impostazione delle Tag per inizio e fondo scala SCADA 1
        Call ScriviTag(CStr(NumeroLinea) & ".AM" & Format(i, "000") & "_ISI", ParametriStrumenti(i).ISI)
        Call ScriviTag(CStr(NumeroLinea) & ".AM" & Format(i, "000") & "_FSI", ParametriStrumenti(i).FSI)
    Next i

    Exit Sub
Gesterrore:
    Call WindasLog("LeggiConfigurazione7AssegnaParametriWinCCScale: " & Error(Err()), 1, OPC)

End Sub
