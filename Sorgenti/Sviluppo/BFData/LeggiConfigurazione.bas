Attribute VB_Name = "LeggiConfigurazione"
Option Explicit

Dim rs As Object

Private Sub LeggiConfigurazione7GenericheLinea()

    Dim indice As Integer

    On Error GoTo GestErrore
    
    '***** generiche per linea *****
    rs.SelectionFast "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & StationCode & "' ORDER BY CC_CODE"
    While (Not (rs.IsEOF))
        indice = rs.getValue("cc_code")
        Generiche(indice).Par = rs.getValue("cc_value")
        Generiche(indice).Testo = rs.getValue("cc_text")
        Generiche(indice).Descrizione = rs.getValue("cc_description")
        
        rs.MoveNext
    Wend
    
    Abilita48H = CBool(Trim(Generiche(i48H).Par))
    AbilitaTrimestre = CBool(Trim(Generiche(iTrimestrale).Par))
    Nome_File_4343 = Trim(Generiche(i4343_File).Testo)
    Nome_Impianto_4343 = Trim(Generiche(i4343_Impianto).Testo)
    Nome_Software_4343 = Trim(Generiche(i4343_SW).Testo)
    
    Exit Sub
GestErrore:

    Call WindasLog("LeggiConfigurazione7GenericheLinea: " & Error(Err()), 1)

End Sub

Sub LeggiConfigurazione7()

    Dim iIdx As Integer
    Dim ii As Integer
    Dim Linea As Integer
    On Error GoTo GestErrore
    
    On Local Error GoTo GestErrore
    
    NewDataObj rs
    
    iIdx = 0
    
    rs.SelectionFast "SELECT was_measures.*,wds_gentab.gt_description, wds_gentab.gt_str2 FROM was_measures inner join wds_gentab on c2=gt_code where gt_type = 'params' AND cm_stationcode = '" & StationCode & "' order by C1"
    
    Do While Not rs.IsEOF
        gaConfigurazioneArchivio(iIdx).STRUM.CodiceParametro = rs.getValue("c1")
        gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro = rs.getValue("c2")
        gaConfigurazioneArchivio(iIdx).STRUM.DescrParametro = rs.getValue("gt_description")
        gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisuraTq = rs.getValue("gt_str2")
        gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisura = rs.getValue("c4")
        gaConfigurazioneArchivio(iIdx).STRUM.NroDecimali = rs.getValue("c5")
        gaConfigurazioneArchivio(iIdx).STRUM.ISE = Val(Replace(rs.getValue("c6"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.FSE = Val(Replace(rs.getValue("c7"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.ISI = Val(Replace(rs.getValue("c8"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.FSI = Val(Replace(rs.getValue("c9"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.FSI2 = Val(Replace(rs.getValue("c10"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAttenzione = Val(Replace(rs.getValue("c11"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAllarme = Val(Replace(rs.getValue("c12"), ",", "."))
        
        'Alby Febbraio 2016 aggiunte soglie attenzione ed allarme giornaliere
        'gestisco le soglie non come % del limite ma come valori assoluti
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAttenzioneGiornaliera = Val(Replace(rs.getValue("c75"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAllarmeGiornaliera = Val(Replace(rs.getValue("c76"), ",", "."))
        'Federica ottobre 2017
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAttenzioneMensile = Val(Replace(rs.getValue("L10"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaAllarmeMensile = Val(Replace(rs.getValue("L11"), ",", "."))
        
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteInferiore = Val(Replace(rs.getValue("c13"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteSuperiore = Val(Replace(rs.getValue("c14"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteInferioreOrario = Val(Replace(rs.getValue("c15"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteSuperioreOrario = Val(Replace(rs.getValue("c16"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.Acquisizione = rs.getValue("c17")
        gaConfigurazioneArchivio(iIdx).STRUM.TipoAcquisizione = rs.getValue("c18")
        gaConfigurazioneArchivio(iIdx).STRUM.OpzioniAcquisizione = rs.getValue("c19")
        gaConfigurazioneArchivio(iIdx).STRUM.TipoStrumento = rs.getValue("c20")
        'Alby Settembre 2014 tolto in quanto aggiornato in BFwinCC e non più compatibile in BFdata NON serve
        'gaConfigurazioneArchivio(iIdx).STRUM.NroMorsetto = rs.GetValue("c21")
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaValidazione = Val(Replace(rs.getValue("c22"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.iddatabase = rs.getValue("c23")
        gaConfigurazioneArchivio(iIdx).STRUM.Elaborazioni = rs.getValue("c24")
        gaConfigurazioneArchivio(iIdx).STRUM.MaxIncremento = Val(Replace(rs.getValue("c25"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.MinEscursione = Val(Replace(rs.getValue("c26"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.MaxEscursione = Val(Replace(rs.getValue("c27"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMediaSemiorariaColonnaA = Val(Replace(rs.getValue("c28"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMediaSemiorariaColonnaB = Val(Replace(rs.getValue("c29"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMediaOraria = Val(Replace(rs.getValue("c30"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMediaGiornaliera = Val(Replace(rs.getValue("c31"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMedia48Ore = Val(Replace(rs.getValue("c32"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteMediaMensile = Val(Replace(rs.getValue("c33"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteFlussoMassaMensile = Val(Replace(rs.getValue("c34"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteFlussoMassaAnnuale = Val(Replace(rs.getValue("c35"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.Invalida = rs.getValue("c36")
        gaConfigurazioneArchivio(iIdx).STRUM.UsaDatoStimato = Val(rs.getValue("c37"))
        gaConfigurazioneArchivio(iIdx).STRUM.ValoreStimato = Val(Replace(rs.getValue("c38"), ",", "."))
        
        'Alby Dicembre 2015
        'gaConfigurazioneArchivio(iIdx).STRUM.DigitaleCambioScala = rs.GetValue("c39")
        
        '*** QAL2
        gaConfigurazioneArchivio(iIdx).STRUM.LimiteRilevabilita = Val(Replace(rs.getValue("c40"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.m = Val(Replace(rs.getValue("c41"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.q = Val(Replace(rs.getValue("c42"), ",", "."))
        'C43 riservato per il Range di taratura
        gaConfigurazioneArchivio(iIdx).STRUM.IntervalloConfidenza = Val(Replace(rs.getValue("c44"), ",", "."))
        
        '*** QAL3
        'C46 riservato per ZERO sams
        'C47 riservato per SPAN sams
        'C48 --> C52 dati per Foglio CUSUM
        
        'C54 disponibile
        
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaMinimaIstantanea = Val(Replace(rs.getValue("c55"), ",", "."))
        gaConfigurazioneArchivio(iIdx).STRUM.SogliaMassimaIstantanea = Val(Replace(rs.getValue("c56"), ",", "."))
        
        gaConfigurazioneArchivio(iIdx).STRUM.ScritturaMisura = Val(rs.getValue("c60"))
        gaConfigurazioneArchivio(iIdx).STRUM.LetturaMisura = Val(rs.getValue("c61"))    '##### letto 2 volte!
        
        '*** DS4343
        gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorIst_TQ = rs.getValue("c65")
        gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorMed_TQ = rs.getValue("c66")
        gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorMed_EL = rs.getValue("c67")
        gaConfigurazioneArchivio(iIdx).STRUM.OrdineScritturaADIADM = IIf(rs.getValue("C68") <> "", rs.getValue("c68"), -9999)

        'michele ottobre 2013 OPC: tag per invio ultime medie orarie a DCS via OPC
        gaConfigurazioneArchivio(iIdx).STRUM.tagOPC_UMO = rs.getValue("c73")
        
        'luca marzo 2017
        gaConfigurazioneArchivio(iIdx).STRUM.QAL2suTQ = IIf(Val(Trim(rs.getValue("c61"))) = 1, True, False)
        'maMisureStrumenti(iIdx).Scansione = ControlloConfigStrumenti(iIdx)
        iIdx = iIdx + 1: rs.MoveNext
    Loop
    
    'Numero parametri configurati
    gnNroParametriStrumenti = iIdx - 1

    Call LeggiConfigurazione7GenericheLinea
    Call LeggiConfigurazione7AssegnaParametri
    
    Set rs = Nothing
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata LeggiConfigurazione7: " + Error(Err), 1)
    'Alby Dicembre 2015
    Resume Next
    
End Sub

Private Sub LeggiConfigurazione7AssegnaParametri()

    On Error GoTo GestErrore
        
    'Ingressi da configurare sempre in configurazione generiche di linea
    'Se non presenti mettere -1
    IngressoIMPIANTO = IIf(CInt(Generiche(iStatoImpianto).Par) = -1, -1, CodParametro(CInt(Generiche(iStatoImpianto).Par)))
    IngressoQFUMI = IIf(CInt(Generiche(iPortata).Par) = -1, -1, CodParametro(CInt(Generiche(iPortata).Par)))
    IngressoTemp = IIf(CInt(Generiche(iTemperatura).Par) = -1, -1, CodParametro(CInt(Generiche(iTemperatura).Par)))
    IngressoPress = IIf(CInt(Generiche(iPressione).Par) = -1, -1, CodParametro(CInt(Generiche(iPressione).Par)))
    IngressoH2O = IIf(CInt(Generiche(iH2O).Par) = -1, -1, CodParametro(CInt(Generiche(iH2O).Par)))
    IngressoO2 = IIf(CInt(Generiche(iO2).Par) = -1, -1, CodParametro(CInt(Generiche(iO2).Par)))
    IngressoO2Umido = IIf(CInt(Generiche(iO2Umido).Par) = -1, -1, CodParametro(CInt(Generiche(iO2Umido).Par)))
    
    'Vedere se serve per il ricalcolo dei parametri
    'IngressoVelocita = IIf(CInt(Par(29)) = -1, -1, CodParametro(CInt(Par(29))))
    
    Exit Sub
    
GestErrore:
    Call WindasLog("BFdata LeggiConfigurazione7AssegnaParametri " + Error(Err), 1)

End Sub

