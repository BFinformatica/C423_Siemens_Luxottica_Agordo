Attribute VB_Name = "Generic"
Option Explicit

Public OPCinErrore As Boolean   'Federica novembre 2017

' Colori
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const GRAY = &HC0C0C0
Global Const BROWN = &H80&
Global Const VIOLET = &H800080
Global Const DARK_RED = &H80&
Global Const DARK_YELLOW = &H8080&
Global Const DARK_GREEN = &H8000&
Global Const DARK_BLUE = &H800000
Global Const DARK_CYAN = &HC0C000
Global Const DARK_GRAY = &H808080

'michele ottobre 2013 OPC
Global Const ULTIMA_MO = 0
Global Const ULTIMA_MG = 1
Global OPC_CONNESSO As Boolean

'Federica gennaio 2018 - Nuova gestione connessioni
Public Type ConnectionsDB
    StationCode As String
    AppServer As String
    AppDatabase As String
    AppDBType As String
    AppDBUser As String
    AppDBPwd As String
    AppScheduleWorking As Boolean
    AppDbVersion As String
    AppRS As Object
    AppOrderSAD As Integer
    AppDefaultDB As Boolean
End Type
Public ConnessioneValida As Boolean
Public connDB() As ConnectionsDB
Public defaultIdxConn As Integer

'******** numero massimo di parametri configurati ********amministratore
Public Const NRO_MAX_PAR_CONFIG = 64

' Strutture per il File di Configurazione: Sezione Strumenti
Type ConfigStrumentiType
    CodiceParametro                 As String * 3   'Codice del Parametro
    NomeParametro                   As String * 10  'Nome della Grandezza
    DescrParametro                  As String * 30  'Descrizione della Grandezza
    UnitaMisura                     As String * 8   'Unità di Misura
    UnitaMisuraTq                    As String * 8   'Unità di Misura talQuale
    NroDecimali                     As Integer      'Nro di Decimali in Visualizzazione
    ISE                             As Double       'Inizio Scala Elettrico
    FSE                             As Double       'Fondo Scala Elettrico
    ISI                             As Double       'Inizio Scala Ingegneristico
    FSI                             As Double       'Fondo Scala Ingegneristico
    FSI2                            As Double       'FSI2 secondo fondo scala ingegneristico
    SogliaAttenzione                As Double       'Soglia di Attenzione
    SogliaAllarme                   As Double       'Soglia di Allarme
    
    'Alby Febbraio 2016
    SogliaAttenzioneGiornaliera     As Double       'Soglia di Attenzione
    SogliaAllarmeGiornaliera        As Double       'Soglia di Allarme
    'Federica ottobre 2017
    SogliaAttenzioneMensile         As Double       'Soglia di Attenzione
    SogliaAllarmeMensile            As Double       'Soglia di Allarme
    
    LimiteInferiore                 As Double       'Limite Inferiore
    LimiteSuperiore                 As Double       'Limite Superiore
    LimiteInferioreOrario           As Double       'Limite Inferiore (4byte)
    LimiteSuperioreOrario           As Double       'Limite Superiore
    Acquisizione                    As Integer      'Acquisizione ON/OFF
    TipoAcquisizione                As Integer      'Tipo di Acquisizione (Comunicazione Analogica, Seriale, ...)
    OpzioniAcquisizione             As String * 30  'Campo a Disposizione per Eventuali Informazioni Aggiuntive (Nome File, Nro Porta Seriale, ..)
    TipoStrumento                   As Integer      'Tipo di Strumento (da NON Acquisire con la Scheda Analogica)
    NroMorsetto                     As Integer      'Numero del Morsetto
    SogliaValidazione               As Integer      'Soglia di Validazione (% Dati Elementari Validi nell'ora)
    iddatabase                      As Integer      'Id database
    Elaborazioni                    As String * 3   'Elaborazioni Particolari (S=Secco;N=Normalizzato;C=Compensato)
    MaxIncremento                   As Double       'Limite Incrementale tra Dati Elementari Consecutivi
    MinEscursione                   As Double       'Min Escursione tra Dati Elementari nell'ora
    MaxEscursione                   As Double       'Max Escursione tra Dati Elementari nell'ora
    
    DigitaleCambioScala             As Integer
    
    '**** per 133
    LimiteMediaSemiorariaColonnaA   As Double
    LimiteMediaSemiorariaColonnaB   As Double
    
    LimiteMediaOraria               As Double
    LimiteMediaGiornaliera          As Double
    LimiteMedia48Ore                As Double
    LimiteMediaMensile              As Double
    LimiteFlussoMassaMensile        As Double
    LimiteFlussoMassaAnnuale        As Double
    
    Invalida                        As String * 250 'allarmi invalidanti
    UsaDatoStimato                  As Integer      'flag che indica se utilizzare o no il valore stimato anzichè la misura acquisita
    ValoreStimato                   As Double       'valore stimato del parametro
    
    LimiteRilevabilita              As Double       'per QAL2
    m                               As Double
    q                               As Double
    IntervalloConfidenza            As Double
    
    SogliaMinimaIstantanea          As Double
    SogliaMassimaIstantanea         As Double
    
    ScritturaMisura                 As Double       'AO di uscita ADAM, registro Modbus per invio al DCS ecc.
    LetturaMisura                   As Double       'registro di lettura Modbus per dati da DCS
    
    CodiceMonitorIst_TQ             As String * 40  'Codice monitor DDS4343 dati istantanei tal quali (file .SAD)
    CodiceMonitorMed_TQ             As String * 40  'Codice monitor DDS4343 dati mediati tal quali (file .MEDIE) codice V
    CodiceMonitorMed_EL             As String * 40  'Codice monitor DDS4343 dati mediati elaborati (file.MEDIE) codice E
    OrdineScritturaADIADM           As Double
    PosRiassuntivo                  As Integer      'Posizione nella pagina riassuntiva del BFLab

    'michele ottobre 2013 OPC: tag per invio ultime medie orarie a DCS via OPC
    tagOPC_UMO                      As String
    
    'luca marzo 2017
    QAL2suTQ                        As Boolean
End Type

Type ConfigurazioneType
    STRUM   As ConfigStrumentiType
    'DIG     As ConfigDigitaleType
    'TAR     As ConfigTaratureType
End Type

Global gaConfigurazioneArchivio(NRO_MAX_PAR_CONFIG) As ConfigurazioneType
Global gnNroParametriStrumenti As Integer
Global StationCode As String
Global CodiceSezione As String * 2
Global CodiceStabilimento As String * 4
Global PathARPA As String
Global PathARPA_FileUnico As String

Global Tabella As String

'*** 4343 ***
Global Nome_Software_4343 As String
Global Nome_File_4343 As String
Global Nome_Impianto_4343 As String
Public Const i4343_File = 100
Public Const i4343_Impianto = 101
Public Const i4343_SW = 102

Global gsDirLavoro As String
Global Abilita48H As Boolean
Global AbilitaTrimestre As Boolean
'Federica gennaio 2018
Type LineaGeneriche
    Par As String
    Testo As String
    Descrizione As String
End Type
Global Generiche(10000) As LineaGeneriche
Public Const iO2RIF = 0
Public Const iStatoImpianto = 20
Public Const iPortata = 21
Public Const iTemperatura = 22
Public Const iPressione = 23
Public Const iH2O = 24
Public Const iO2 = 25
Public Const iO2Umido = 26
Public Const iDivisorePerFlussi = 56
Public Const i48H = 57
Public Const iTrimestrale = 58

Global Const CONF_ELAB_SECCO = "Secco"
Global Const CONF_ELAB_NORM = "Normalizzazione"
Global Const CONF_ELAB_COMP = "Compensazione"

'**** michele - gestione stato impianto *****
'Global StatoImpianto(23, 10, 60) As Integer
Global PercRegime(23, 65) As Single
Global PercSpegnimento(23, 65) As Single
Global PercManutenzione(23, 65) As Single
'Global PercTransitorio(23, 60) As Single
Global PercFermo(23, 65) As Single
Global PercMinTec(23, 65) As Single
Global PercGuasto(23, 65) As Single
Global PercAnomalo(23, 65) As Single
'luca gennaio 2015 gestione stato impianto 37
Global PercPolveri(23, 65) As Single
Global PercAltro(23, 65) As Single  'Federica marzo 2018 stato impianto 38
Global statoimp(23, 65) As Single

'**** gestione elaborazioni
Global IngressoO2 As Integer
Global IngressoO2Umido As Integer   'Federica settembre 2017
Global IngressoH2O As Integer
Global IngressoPress As Integer
Global IngressoTemp As Integer
Global IngressoQFUMI As Integer
Global IngressoIMPIANTO As Integer
Global IngressoCO As Integer
Global IngressoCOH As Integer
Global IngressoCOL As Integer

Global ContaSecondiMediaOra(23, NRO_MAX_PAR_CONFIG, 720) As Integer
Global PienamenteOperativo(23, 720, 2) As Integer
Global Valore_5_Secondi(23, NRO_MAX_PAR_CONFIG, 720) As Double
Global Valore_5_Secondi_N(23, NRO_MAX_PAR_CONFIG, 720) As Double
Global Status_5_Secondi(23, NRO_MAX_PAR_CONFIG, 720) As String
Global Status_5_Secondi_N(23, NRO_MAX_PAR_CONFIG, 720) As String
'Alby Luglio 2013 Enipower Bolgiano aggiunta dati stimati
Global Valore_5_Secondi_S(23, NRO_MAX_PAR_CONFIG, 720) As Double
Global Status_5_Secondi_S(23, NRO_MAX_PAR_CONFIG, 720) As String
'dati stimati elaborati
Global Valore_5_Secondi_SN(23, NRO_MAX_PAR_CONFIG, 720) As Double
Global Status_5_Secondi_SN(23, NRO_MAX_PAR_CONFIG, 720) As String

Global ValIstPerScarto(23, NRO_MAX_PAR_CONFIG, 720, 2) As Double
Global ContaTuttiSecondiMediaOra(23, NRO_MAX_PAR_CONFIG, 720) As Integer
'daniele luglio 2013 bolgiano: aumento da 60 a 65
Global ContaTutti_5_secondi(23, NRO_MAX_PAR_CONFIG, 2, 65) As Integer
Global ContaOraOK(23, NRO_MAX_PAR_CONFIG, 2, 65) As Integer
Global ContaOraOK_AUX(23, NRO_MAX_PAR_CONFIG, 2, 65) As Integer
Global MedieOra(23, NRO_MAX_PAR_CONFIG, 2, 65) As Double
Global StsMedieOra(23, NRO_MAX_PAR_CONFIG, 2, 65) As String
'Global NumeroDatiPerMedia(23, NRO_MAX_PAR_CONFIG) As Integer
Global massimo(23, NRO_MAX_PAR_CONFIG, 2, 65) As Double
Global minimo(23, NRO_MAX_PAR_CONFIG, 2, 65) As Double
'Global flusso_massa(23, NRO_MAX_PAR_CONFIG) As Double
Global StdDev(23, NRO_MAX_PAR_CONFIG, 2, 65) As Double
'Global NumeroTotaleDati(23, NRO_MAX_PAR_CONFIG) As Integer
'Global CtrTotAcquisizioni(23, NRO_MAX_PAR_CONFIG) As Integer
'Global StsMedieOraImpianto(23) As String * 2
Global DatoFlussoMassa(23, NRO_MAX_PAR_CONFIG, 2, 65) As Double
Global DatiDaElaborare(5) As Boolean
Global StrLabel As String

'Alby Gennaio 2016
Global DatiPerWinCC As Boolean

'Alby Luglio 2013 Enipower Bolgiano
'spostate le dichiarazione a livello di modulo
'in modo da poter strutturare il software
'per renderlo leggibile!!!!
Global Elabdate As Date
Public nn As Integer
Public nl As Integer
Public Ora As Integer
Public TipoMedia As Integer
Public MaxDati As Integer
Public secondi As Integer
Public SommatoriaOra As Double
Public numeromedie(23) As Integer
Public nmedie As Integer
Public xx As Integer
Public yy As Integer
Public SumDevStd(2) As Double
Public StatiImpianto(10, 1) As Single
Public OraFine As Integer
Public rs As Object
Public strSQL As String
Public Stato_Monitor(10, 1) As Integer
Public sts_n As Integer
Public SommatoriaOra_AUX As Double

'daniele luglio 2013 bolgiano: aggiungo variabile per eventuali dati stimati sul sad dei valori misurati (vedi q metano)
'daniele settembre 2013 bolgiano: correggo gestione Q_Gas ausiliario
'Public intDatiAuxOK(NRO_MAX_PAR_CONFIG) As Integer
Public intDatiAuxOK(24, NRO_MAX_PAR_CONFIG) As Integer

'Alby Ottobre 2013 dichiarazioni ereditate da WinCC .....
Public ValIst(3, 72), OldValIst(72), ValIstBreeze(72), Status(3, 72)

'michele ottobre 2013: per colonna con l'ossigeno di riferimento
Global O2riferimento As Single

'Alby Dicembre 2015
Global UltimaMediaOraria As Double
Global MediaInCorsoGiorno As Double

Public UltimaMediaGiorno As Double
Public UltimaMedia48h As Double
Public MediaInCorso48h As Double
Public NrDatiInCorso48h As Double
'luca 06/09/2016 ID medie 48H
Public IDUltima48H As Double
Public IDCostruzione48H As Double
'luca 16/09/2016 status 48H
Public StatusUltima48H As String
Public StatusCostruzione48H As String

Global Ruolo As String
Global Client As Boolean

Public Const strInTransitorio = "'30', '31', '32'"  'Federica dicembre 2017
Public Const strValidValidflags = "'VAL','AUX'" '### Da leggere da database
'Federica dicembre 2017 - Tag WinCC Medie
Public TagUltimaMedia As String
Public TagValiditaUltimaMedia As String
Public TagVisualizzazioneUltimaMedia As String
Public TagVisualizzazioneValiditaUltimaMedia As String
Public TagIDUltimaMedia As String
Public TagMediaCorrente As String
Public TagValiditaMediaCorrente As String
Public TagVisualizzazioneMediaCorrente As String
Public TagVisualizzazioneValiditaMediaCorrente As String
Public TagIDMediaCorrente As String
Public TagMediaPrevisionale As String
Public TagVisualizzazioneMediaPrevisionale As String
Public InizioTag As String

Global Dati(18000, 100, 1) As String


    




'luca 06/09/2016 revisiono funzione utilizzando codifiche numeriche
Function DeterminaColore(valore, indice, tipo) As Integer

    Dim Soglia(1) As Double

    'Alby Febbraio 2016
    On Error GoTo GestErrore
    
    'di defalut colore verde
    'luca 06/09/2016 colore default 0
    DeterminaColore = 0
    
    If valore = -9999 Then
        'luca 06/09
        DeterminaColore = 1
        Exit Function
    End If
    
    'se il dato è valido controllo le soglie ed eventualmente coloro
    If UCase(Tabella) <> "WDS_10MINCO" Then
        Select Case tipo
            Case 0      'su limiti medi orari **********************************************************************************
                
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzione > 0 Then
                    'luca 06/09/2016 soglia attenzione -> giallo
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzione Then DeterminaColore = 5
                End If
                
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAllarme > 0 Then
                    'luca 06/09/2016 soglia allarme -> rosso
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAllarme Then DeterminaColore = 2
                End If
            Case 1      'su limiti medi giornalieri ****************************************************************************
                
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzioneGiornaliera > 0 Then
                    'luca 06/09/2016 soglia attenzione -> giallo
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzioneGiornaliera Then DeterminaColore = 5
                End If
                
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAllarmeGiornaliera > 0 Then
                    'luca 06/09/2016 soglia allarme -> rosso
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAllarmeGiornaliera Then DeterminaColore = 2
                End If
                
            Case 2      'su limiti medi mensili ****************************************************************************
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzioneMensile > 0 Then
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAttenzioneMensile Then DeterminaColore = 5
                End If
                
                If gaConfigurazioneArchivio(indice).STRUM.SogliaAllarmeMensile > 0 Then
                    If valore > gaConfigurazioneArchivio(indice).STRUM.SogliaAllarmeMensile Then DeterminaColore = 2
                End If
        End Select
    End If

    Exit Function

GestErrore:
    Call WindasLog("DeterminaColore " + Error(Err), 1)

End Function

Sub ElaboraAggiornaSQL(Tabella, Data, iIdx1, Media, Status, Somma, SommaTot, ContaValidiInMarcia, ContaInMarcia)

    Dim rsDati2 As Object
    Dim strSQL As String

    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    NewDataObj rsDati2

    strSQL = "DELETE FROM " + Tabella + " WHERE DT_stationcode='" + StationCode + "' and dt_measurecod='" + Trim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro) + "' and dt_date='" + Data + "'"
    rsDati2.ExecuteSQL (strSQL)


    strSQL = "INSERT INTO " + Tabella + " (dt_stationcode,dt_measurecod,dt_date,dt_value,dt_validflag, dt_fm, dt_fmtot,dt_nr, dt_nrtot)"
    strSQL = strSQL + " VALUES ('" + StationCode + "','" + Trim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro) + "','" + Data + "',"
    strSQL = strSQL + Trim(Str(Media)) + ",'" + Status + "'," + Trim(Str(Somma)) + "," + Trim(Str(SommaTot)) + "," + Trim(Str(ContaValidiInMarcia)) + "," + Trim(Str(ContaInMarcia)) + ")"
    rsDati2.ExecuteSQL strSQL

    Set rsDati2 = Nothing

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaSQL: " + Error(Err), 1)

End Sub

'luca 22/07/2016 aggiungo tag colore
Sub ElaboraAggiornaWinCCTag(iIdx1, TagValore, TagColore, valore, tipo)

    Dim ValTagColore As String

    'Alby Dicembre 2015
    On Error GoTo GestErrore

    'Alby Febbraio 2016
    'luca 22/07/2016
    ScriviTag TagValore, valore
    'luca 22/07/2016 uso il codiceparametro (posizione BFLab)
    ValTagColore = DeterminaColore(valore, iIdx1, tipo)
    ScriviTag TagColore, ValTagColore

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaWinCCTag: " + Error(Err), 1)
    
End Sub

Sub ElaboraSalvaDatiConcludoSemiora()

    On Error GoTo GestErrore

    
    '***** salvataggio dei dati nel DB *****
    For Ora = 0 To OraFine
        For nmedie = 1 To numeromedie(Ora)
            Form1.Label1.Caption = StrLabel & " Ora:" & Str(Ora) & " - Media:" & Str(nmedie)
            For nn = 0 To gnNroParametriStrumenti
                DoEvents
                If Not Client Then
                    Call ElaboraSalvaDatiSQL(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
                End If
            Next nn
        Next nmedie
    Next Ora
                
    'Alby Dicembre 2015 se è in automatico da BFwincc e se l'ora è la precedente del giorno corrente
    If InStr(Command, "auto") > 0 And (Format(Elabdate, "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") And Format(Ora, "00") = Format(DateAdd("h", -1, Now), "hh")) Then
        'Elaboro dati e aggiorno WinCC
        DatiPerWinCC = True
        Call InizializzaWinCC
        Call ElaboraSalvaDatiMedieNF("48", nn, Elabdate)
        Call ElaboraAggiornaMedia(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
    Else
        DatiPerWinCC = False
        If Ora = 23 Then
            'se rielabora una giornata che non è il giorno corrente NON aggiorna WinCC
            'e elabora medie 48h, giornaliere e mensili solo all'ultima ore nel ciclo la 23 (0-23)
            Call ElaboraSalvaDatiMedieNF("48", nn, Elabdate)
            Call ElaboraAggiornaMedia(Ora, TipoMedia, nn, Elabdate, 0, nmedie)
        End If
    End If

    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiConcludoSemiora " + Error(Err), 1)


End Sub

Sub ElaboraSalvaDatiFlussiMassa()

    Dim nn As Integer
    Dim DatoSostitutivo As Double
    
    On Error GoTo GestErrore
    
    'Federica novembre 2017 - Gestione mancanza parametro
    If (CDbl(Generiche(iDivisorePerFlussi).Par) < 1) Then
        Call WindasLog("ElaboraSalvaDatiFlussiMassa: Divisore per calcolo flussi non impostato!", 1)
        Exit Sub
    End If

    '***** calcolo flussi di massa con i dati elaborati ******
    For nn = 0 To gnNroParametriStrumenti
        If IngressoQFUMI > -1 Then
            For Ora = 0 To OraFine
                For nmedie = 1 To numeromedie(Ora)
                    'luca luglio 2017 gestisco anche il VAH come dato valido
                    'If StsMedieOra(Ora, IngressoQFUMI(nl), 1, nmedie) = "VAL" Then
                    If InStr("VAL VAH", StsMedieOra(Ora, IngressoQFUMI, 1, nmedie)) > 0 Then
                        'luca luglio 2017 gestisco anche il VAH come dato valido
                        'If StsMedieOra(Ora, nn, 1, nmedie) = "VAL" Then
                        If InStr("VAL VAH", StsMedieOra(Ora, nn, 1, nmedie)) > 0 Then
                            If nn = IngressoQFUMI Then
                                'Riscrivo la portata
                                DatoFlussoMassa(Ora, nn, 1, nmedie) = MedieOra(Ora, IngressoQFUMI, 1, nmedie)
                            Else
                                DatoFlussoMassa(Ora, nn, 1, nmedie) = MedieOra(Ora, nn, 1, nmedie) * MedieOra(Ora, IngressoQFUMI, 1, nmedie) / CDbl(Generiche(iDivisorePerFlussi).Par)
                            End If
                        Else
                            DatoFlussoMassa(Ora, nn, 1, nmedie) = -9999
                        End If
                    Else
                        'Alby Giugno 2016 se non c'è la portata pongo DatoFlussoMassa a zero
                        DatoFlussoMassa(Ora, nn, 1, nmedie) = -9999
                    End If
                Next nmedie
            Next Ora
        End If
    Next nn
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiFlussiMassa " + Error(Err), 1)
    
End Sub

Sub GestioneQAL2()
    
    Dim i As Integer
    Dim DatoQAL2 As Double
    
    For i = 0 To gnNroParametriStrumenti
        'm
        DatoQAL2 = CaricaDatiQAL2(Elabdate, "C41", i)
        If DatoQAL2 <> -9999 Then gaConfigurazioneArchivio(i).STRUM.m = DatoQAL2
        'q
        DatoQAL2 = CaricaDatiQAL2(Elabdate, "C42", i)
        If DatoQAL2 <> -9999 Then gaConfigurazioneArchivio(i).STRUM.q = DatoQAL2
        'IC
        DatoQAL2 = CaricaDatiQAL2(Elabdate, "C44", i)
        If DatoQAL2 <> -9999 Then gaConfigurazioneArchivio(i).STRUM.IntervalloConfidenza = DatoQAL2
        'Limite rilevabilità
        DatoQAL2 = CaricaDatiQAL2(Elabdate, "C40", i)
        If DatoQAL2 <> -9999 Then gaConfigurazioneArchivio(i).STRUM.LimiteRilevabilita = DatoQAL2
    Next i
   
    Exit Sub
    
GestErrore:
    Call WindasLog("GestioneQAL2 " + Error(Err), 1)

End Sub

'luca aprile 2017
Function CaricaDatiQAL2(Elabdate As Date, ColumnField As String, iIdx As Integer) As Double

    Dim rsDati As Object
    Dim strSQL As String

    On Error GoTo GestErrore

    CaricaDatiQAL2 = -9999
    NewDataObj rsDati

    With rsDati
        strSQL = "SELECT * FROM WLS_CFGLOG WHERE station='" + StationCode + "' AND Parameter='" + Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro) + "'"
        strSQL = strSQL + " AND ColumnField='" + ColumnField + "' AND date<='" + Format(Elabdate, "yyyymmdd") + "' ORDER BY date,Time"
        .SelectionFast strSQL
        If Not .iseof Then
            .movelast
            CaricaDatiQAL2 = Val(Replace(.GetValue("NewValue"), ",", "."))
        End If
    End With
    
    Set rsDati = Nothing
    
    Exit Function
    
GestErrore:
    Call WindasLog("BFData CaricaDatiQAL2 " + Error(Err), 1)
    
End Function
Sub ElaboraSalvaDatiInizializzo()
 
    On Error GoTo GestErrore
    

    If UCase(Tabella) = "WDS_HALF" Then
        MaxDati = 360: TipoMedia = 0
        For Ora = 0 To 23
            numeromedie(Ora) = 2
        Next Ora
    ElseIf UCase(Tabella) = "WDS_10MINCO" Then
        MaxDati = 120: TipoMedia = 1
        For Ora = 0 To 23
            numeromedie(Ora) = 6
        Next Ora
    ElseIf UCase(Tabella) = "WDS_ELAB" Then
        MaxDati = 720: TipoMedia = 2
        For Ora = 0 To 23
            numeromedie(Ora) = 1
        Next Ora
    ElseIf UCase(Tabella) = "WDS_AUTO" Then
        MaxDati = 12: TipoMedia = 3
        For Ora = 0 To 23
            numeromedie(Ora) = 60
        Next Ora
    End If
    
    If InStr(UCase(Command), "AUTO") > 0 Then
    
        '***** automatico lanciato da BFLab *****
        If hour(Elabdate) = 0 Then
            OraFine = 23
        Else
        
            Select Case TipoMedia
                
                Case 0
                    '***** media semioraria *****
                    If minute(Elabdate) < 30 Then
                        OraFine = hour(Elabdate) - 1
                        'numeromedie = 2
                    Else
                        OraFine = hour(Elabdate)
                        numeromedie(OraFine) = 1
                    End If
                    
                Case 1
                    '***** media 10 minuti CO *****
                    If minute(Elabdate) < 10 Then
                        OraFine = hour(Elabdate) - 1
                        'numeromedie = 6
                    Else
                        OraFine = hour(Elabdate)
                        numeromedie(OraFine) = Int(minute(Elabdate) / 10)
                    End If
                
                Case 2
                    '***** media oraria *****
                    OraFine = hour(Elabdate) - 1
                    'numeromedie = 1
            
            End Select
        
        End If
    
    Else
    
        '***** manuale ******
        If Date = Elabdate Then
            OraFine = hour(Now) - 1
        Else
            OraFine = 23
        End If
        
    End If
    Exit Sub
       
GestErrore:
    Call WindasLog("ElaboraSalvaDatiInizializzo " + Error(Err), 1)
    Resume fine:
fine:
    
End Sub
'luca aprile 2017
Function CalcolaQAL2(indice As Integer, valore As Double) As Double

Dim temp As Double

On Error GoTo GestErrore

    temp = valore
    
    If Not gaConfigurazioneArchivio(indice).STRUM.QAL2suTQ Then
        'se diverso da -9999 applica QAL2 altrimenti restituisce -9999
        If temp <> -9999 Then
            'QAL2
            If gaConfigurazioneArchivio(indice).STRUM.m <> 0 Then
                temp = temp * gaConfigurazioneArchivio(indice).STRUM.m + gaConfigurazioneArchivio(indice).STRUM.q
            End If
            
            'Intervallo di confidenza *****
            If gaConfigurazioneArchivio(indice).STRUM.IntervalloConfidenza <> 0 Then
                temp = temp - gaConfigurazioneArchivio(indice).STRUM.IntervalloConfidenza
            End If
            
            '***** Limite di rilevabilita *****
            If gaConfigurazioneArchivio(indice).STRUM.LimiteRilevabilita >= 0 Then
                If temp < gaConfigurazioneArchivio(indice).STRUM.LimiteRilevabilita Then
                    temp = gaConfigurazioneArchivio(indice).STRUM.LimiteRilevabilita
                End If
            End If
        End If
    End If
    
    CalcolaQAL2 = temp
    
Exit Function

GestErrore:
Call WindasLog("CalcolaQAL2: " + Error(Err), 1)

End Function

Sub ElaboraSalvaDatiNormalizzaIstantaneo(CheDato As Integer)

    Dim ValoreTQ As Double
    Dim H2O As Double
    Dim O2 As Double
    Dim T As Double
    Dim P As Double
    Dim Status As String
    'luca aprile 2017
    Dim tempH2O As Double
    Dim tempO2 As Double
    
    On Error GoTo GestErrore

    'Alby Luglio 2013 Enipower Bolgiano
    If CheDato = 0 Then
        ValoreTQ = Valore_5_Secondi(Ora, nn, secondi)
        'Alby Novembre 2014
        If IngressoH2O <> -1 Then
            H2O = Valore_5_Secondi(Ora, IngressoH2O, secondi)
            'luca aprile 2017
            tempH2O = CalcolaQAL2(IngressoH2O, H2O)
            H2O = tempH2O
        Else
            H2O = -9999
        End If
        'daniele luglio 2013 bolgiano: richiamava sempre istantaneo O2 grezzo misurato per compensazione ossigeno: se dato misurato è invalido, prendo l'aux
        If IngressoO2 <> -1 Then
            O2 = Valore_5_Secondi(Ora, IngressoO2, secondi)
        Else
            O2 = -9999
        End If
        'luca aprile 2017
        tempO2 = CalcolaQAL2(IngressoO2, O2)
        O2 = tempO2
        If IngressoTemp <> -1 Then
            T = Valore_5_Secondi(Ora, IngressoTemp, secondi)
        Else
            T = -9999
        End If
        If IngressoPress <> -1 Then
            P = Valore_5_Secondi(Ora, IngressoPress, secondi)
        Else
            P = -9999
        End If
        Status = Status_5_Secondi(Ora, nn, secondi)
        'If ValoreTQ <> -9999 Then Stop
        Valore_5_Secondi_N(Ora, nn, secondi) = ElaborazioniDiLegge(ValoreTQ, H2O, O2, T, P, nn, Status)
        Status_5_Secondi_N(Ora, nn, secondi) = Status
    Else
        'daniele luglio 2013 bolgiano: di qui non passa mai!
        ValoreTQ = Valore_5_Secondi_S(Ora, nn, secondi)
        H2O = Valore_5_Secondi_S(Ora, IngressoH2O, secondi)
        'luca aprile 2017
        tempH2O = CalcolaQAL2(IngressoH2O, H2O)
        H2O = tempH2O
        O2 = Valore_5_Secondi_S(Ora, IngressoO2, secondi)
        'luca aprile 2017
        tempO2 = CalcolaQAL2(IngressoO2, O2)
        O2 = tempO2
        T = Valore_5_Secondi_S(Ora, IngressoTemp, secondi)
        P = Valore_5_Secondi_S(Ora, IngressoPress, secondi)
        Status = Status_5_Secondi_S(Ora, nn, secondi)
        'If ValoreTQ <> -9999 Then Stop
        Valore_5_Secondi_SN(Ora, nn, secondi) = ElaborazioniDiLegge(ValoreTQ, H2O, O2, T, P, nn, Status)
        Status_5_Secondi_SN(Ora, nn, secondi) = Status
    End If
    
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiNormalizzaIstantaneo " + Error(Err), 1)
    Resume Next


End Sub

Function EstraiDatoSostitutivo(Misura, Elabdate, Ora)
    
    Dim rsDati As Object
    Dim strSQL As String
    Dim UltimiDati(1 To 6) As Double
    Dim ContaDati As Integer
    Dim ChiaveMisura As String
    Dim NrOre As Integer
    Dim ValoreMassimo As Double
    Dim DataDB As Date
    Dim iIdx As Integer
    
    'Alby Gennaio 2016
    On Error GoTo GestErrore
    
    EstraiDatoSostitutivo = -9999
    NewDataObj rsDati
    ChiaveMisura = Trim(gaConfigurazioneArchivio(Misura).STRUM.NomeParametro)
    
    DataDB = Format(Elabdate, "dd/mm/yyyy") + " " + Format(Ora, "00") + ".00"
    Do
        strSQL = "SELECT * FROM WDS_ELAB WHERE dt_stationcode='" + StationCode + "' AND dt_measurecod='" + ChiaveMisura + "' AND dt_date='" + Format(DataDB, "yyyymmdd") + "'"
        strSQL = strSQL + " AND dt_hour='" + Format(DataDB, "hh.nn") + "'"
        rsDati.SelectionFast strSQL
        If Not rsDati.iseof Then
            If rsDati.GetValue("dt_validflag") = "VAL" Then
                ContaDati = ContaDati + 1
                UltimiDati(ContaDati) = rsDati.GetValue("dt_value")
            End If
        End If
        NrOre = NrOre + 1
        DataDB = DateAdd("h", -1, DataDB)
        If ContaDati = 6 Then Exit Do
        If NrOre > 100 Then
            'Call WindasLog("Raggiunte 100 ore nella ricerca di un dato valido", 0)
            Exit Function
        End If
    Loop
    
    ValoreMassimo = -9999
    For iIdx = 1 To 6
        If UltimiDati(iIdx) > ValoreMassimo Then ValoreMassimo = UltimiDati(iIdx)
    Next iIdx
    
    EstraiDatoSostitutivo = ValoreMassimo
    
    Set rsDati = Nothing
    
    Exit Function
    
GestErrore:
    Call WindasLog("EstraiDatoSostitutivo " + Error(Err), 1)
    
End Function

Function FormattaNumero(Numero, iIdx1) As String
    
    Dim NrDecimali As Integer

    On Error GoTo GestErrore

    'Alby Luglio 2013 Enipower Bolgiano
    'per formattare numeri
    'se passo alla funziona valore negativo il valore viene presentato con quei decimali
    'se passo alla funzione valore positivo il valore viene presentato con i decimali configurati in Bfdesk
    If iIdx1 < 0 Then
        NrDecimali = Abs(iIdx1)
    Else
        NrDecimali = gaConfigurazioneArchivio(iIdx1).STRUM.NroDecimali
    End If
     
    Select Case NrDecimali
        Case 0
            'daniele luglio 2013 bolgiano: sistemo la formattazione numerica
            FormattaNumero = Replace(CStr(Format(Numero, "0")), ",", ".")
        Case 1
            FormattaNumero = Replace(CStr(Format(Numero, "0.0")), ",", ".")
        Case 2
            FormattaNumero = Replace(CStr(Format(Numero, "0.00")), ",", ".")
        Case 3
            FormattaNumero = Replace(CStr(Format(Numero, "0.000")), ",", ".")
        Case 4
            FormattaNumero = Replace(CStr(Format(Numero, "0.0000")), ",", ".")
        Case 5
            FormattaNumero = Replace(CStr(Format(Numero, "0.00000")), ",", ".")
        Case Else
            FormattaNumero = Replace(CStr(Format(Numero, "0.000000")), ",", ".")
    End Select
    
    
    
    Exit Function
    
GestErrore:
    Call WindasLog("FormattaNumero " + Error(Err), 1)
    Resume Next


End Function


Function GiorniMese(Data) As Integer

    Dim ElaboroData As Date

    'Alby Dicembre 2015
    On Error GoTo GestErrore

    'vado al prossimo mese
    ElaboroData = DateAdd("m", 1, Data)
    
    'mi posiziono al primo giorno di quel mese
    'luca 16/09/2016 formato data inglese non funziona
    'ElaboroData = "01/" + Format(ElaboroData, "mm/yyyy")
    ElaboroData = DateSerial(year(ElaboroData), month(ElaboroData), 1)
    
    'vado al giorno precedente che è l'ultimo giorno di questo mese
    ElaboroData = DateAdd("d", -1, ElaboroData)
    
    'luca 16/09/2016
    'GiorniMese = Val(Format(ElaboroData, "dd"))
    GiorniMese = day(ElaboroData)
    
    Exit Function

GestErrore:
    Call WindasLog("BFdata GiorniMese: " + Error(Err), 1)

End Function

Public Sub QuickSort(arr As Variant, Optional numEls As Variant, Optional descending As Boolean)

    Dim Value As Variant, temp As Variant
    Dim cod_value As Variant, cod_temp As Variant
    Dim sp As Integer
    Dim leftStk(32) As Long, rightStk(32) As Long
    Dim leftNdx As Long, rightNdx As Long
    Dim i As Long, j As Long

    ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)
    ' init pointers
    leftNdx = LBound(arr)
    rightNdx = numEls
    ' init stack
    sp = 1
    leftStk(sp) = leftNdx
    rightStk(sp) = rightNdx

    Do
        If rightNdx > leftNdx Then
            Value = arr(rightNdx, 0)
            i = leftNdx - 1
            j = rightNdx
            ' find the pivot item
            If descending Then
                Do
                    Do: i = i + 1: Loop Until arr(i, 0) <= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j, 0) >= Value
                    temp = arr(i, 0)
                    arr(i, 0) = arr(j, 0)
                    arr(j, 0) = temp
                    
                    cod_temp = arr(i, 1)
                    arr(i, 1) = arr(j, 1)
                    arr(j, 1) = cod_temp
                    
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until arr(i, 0) >= Value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j, 0) <= Value
                    temp = arr(i, 0)
                    arr(i, 0) = arr(j, 0)
                    arr(j, 0) = temp
                    
                    cod_temp = arr(i, 1)
                    arr(i, 1) = arr(j, 1)
                    arr(j, 1) = cod_temp

                Loop Until j <= i
            End If
            ' swap found items
            temp = arr(j, 0)
            arr(j, 0) = arr(i, 0)
            arr(i, 0) = arr(rightNdx, 0)
            arr(rightNdx, 0) = temp
            
            cod_temp = arr(j, 1)
            arr(j, 1) = arr(i, 1)
            arr(i, 1) = arr(rightNdx, 1)
            arr(rightNdx, 1) = cod_temp

            ' push on the stack the pair of pointers that differ most
            sp = sp + 1
            If (i - leftNdx) > (rightNdx - i) Then
                leftStk(sp) = leftNdx
                rightStk(sp) = i - 1
                leftNdx = i + 1
            Else
                leftStk(sp) = i + 1
                rightStk(sp) = rightNdx
                rightNdx = i - 1
            End If
        Else
            ' pop a new pair of pointers off the stacks
            leftNdx = leftStk(sp)
            rightNdx = rightStk(sp)
            sp = sp - 1
            If sp = 0 Then Exit Do
        End If
    Loop
    
    Exit Sub
    
GestErr:
    Debug.Print Now, " QuickSort: " & Err.Description
    Call WindasLog("BFdata QuickSort: " + Error(Err), 1)
    Resume Next
    
End Sub

Sub ElaboraSalvaDatiSQL(periodo As Integer, tipo As Integer, iIdx1 As Integer, Elabdate, tipodato As Integer, numMedia As Integer)

    Dim rsDati As Object
    Dim rs As Object
    Dim strSQL As String
    Dim orarioStr As String
    Dim valore As Double
    Dim validflag As String
    Dim nContaValori As Integer
    Dim stsImpianto As String
    Dim UltimaMedia As Integer
    'Dim data As Date
    Dim NumeroLinea As Integer
    
    On Error GoTo GestErrore
    
    Select Case tipo
    
        Case 0
            '***** media mezzora *****
            If numMedia = 1 Then
                orarioStr = Format(periodo, "00") & ".00"
            ElseIf numMedia = 2 Then
                orarioStr = Format(periodo, "00") & ".30"
            End If
    
        Case 1
            '***** media 10 minuti CO *****
            orarioStr = Format(periodo, "00") & "." & Format(numMedia * 10 - 10, "00")
        
        Case 2
            '***** media oraria *****
            orarioStr = Format(periodo, "00") & ".00"
            
        Case 3
            '***** media del minuto *****
            orarioStr = Format(periodo, "00") & "." & Format(numMedia - 1, "00")
    
        Case Else
            Exit Sub
    
    End Select
    
    NewDataObj rsDati
    
    If ContaTutti_5_secondi(periodo, iIdx1, 0, numMedia) > 0 Then
    
        'Alby Dicembre 2015
        NumeroLinea = 1
        stsImpianto = Trim(Str(statoimp(periodo, numMedia)))
    
        With rsDati

            '***** Lettura record *****
            strSQL = "SELECT DT_VALUE,DT_HOUR,DT_VALIDFLAG,DT_NR,DT_MAX,DT_MIN FROM " & Tabella & " WHERE DT_STATIONCODE = '" & StationCode & "'AND DT_MEASURECOD = '" & gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro & "' AND DT_DATE = " & .ParSQLDate(Elabdate) & " AND DT_HOUR='" & orarioStr & "'"
            If (.SelectionFast(strSQL)) Then

                'aggiorno record preesistente
                'luca 08/11/2016 modifico (gestisco indice per COL - COH)
                strSQL = "UPDATE " & Tabella & " SET " & _
                    "DT_Value = " & Replace(Str(MedieOra(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                    "DT_ValidFlag = " & .ParSQLStr(StsMedieOra(periodo, iIdx1, 1, numMedia)) & "," & _
                    "DT_ValueTQ = " & Replace(Str(MedieOra(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                    "DT_ValidFlag_TQ = " & .ParSQLStr(StsMedieOra(periodo, iIdx1, 0, numMedia)) & "," & _
                    "DT_Custom1 = " & .ParSQLStr(stsImpianto) & "," & _
                    "DT_UnitMeasure = " & .ParSQLStr(Trim(gaConfigurazioneArchivio(iIdx1).STRUM.UnitaMisura)) & "," & _
                    "DT_Nr = " & ContaOraOK(periodo, iIdx1, 1, numMedia) & "," & _
                    "DT_NrTot = " & ContaTutti_5_secondi(periodo, iIdx1, 0, numMedia) & "," & _
                    "DT_StdDev = " & Replace(Str(StdDev(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                    "DT_Min = " & Replace(Str(minimo(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                    "DT_Max = " & Replace(Str(massimo(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                    "DT_Min_TQ = " & Replace(Str(minimo(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                    "DT_Max_TQ = " & Replace(Str(massimo(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                    "DT_Nr_TQ = " & ContaOraOK(periodo, iIdx1, 0, numMedia) & "," & _
                    "DT_StdDev_TQ = " & Replace(Str(StdDev(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                    "DT_FM = " & Replace(Str(DatoFlussoMassa(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                    "DT_Parameter= " & .ParSQLStr("1" & Trim(Str$(gaConfigurazioneArchivio(iIdx1).STRUM.iddatabase))) & _
                    " WHERE DT_StationCode = " & .ParSQLStr(Trim(StationCode)) & " AND " & _
                    "DT_MeasureCod = " & .ParSQLStr(Trim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)) & " AND " & _
                    "DT_Date = " & .ParSQLDate(Elabdate) & " AND " & _
                    "DT_Hour =" & .ParSQLStr(orarioStr)
                .ExecuteSQL (strSQL)
            Else
                'inserisco record
                'NB l'ultimo campo in dt_custom1 è lo stato impianto preso da StsMedieOra(periodo, (indice totale parametri + 1) è sempre lo stato impianto)
                strSQL = "INSERT INTO " & Tabella & " (DT_StationCode,DT_MeasureCod,DT_Date,DT_Hour,DT_Value,DT_ValidFlag,DT_UnitMeasure,DT_Nr,DT_Nr_TQ,DT_Min,DT_Max,DT_Min_TQ,DT_Max_TQ,DT_ValueTQ,DT_ValidFlag_TQ,DT_NrTot,DT_StdDev,DT_StdDev_TQ,DT_FM,DT_Parameter, dt_custom1) VALUES (" & _
                  .ParSQLStr(Trim(StationCode)) & "," & _
                  .ParSQLStr(Trim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)) & "," & _
                  .ParSQLDate(Elabdate) & "," & _
                  .ParSQLStr(orarioStr) & "," & _
                  Replace(Str(MedieOra(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                  .ParSQLStr(StsMedieOra(periodo, iIdx1, 1, numMedia)) & "," & _
                  .ParSQLStr(Trim(gaConfigurazioneArchivio(iIdx1).STRUM.UnitaMisura)) & "," & _
                  ContaOraOK(periodo, iIdx1, 1, numMedia) & "," & _
                  ContaOraOK(periodo, iIdx1, 0, numMedia) & "," & _
                  Replace(Str(minimo(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                  Replace(Str(massimo(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                  Replace(Str(minimo(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                  Replace(Str(massimo(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                  Replace(Str(MedieOra(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                  .ParSQLStr(StsMedieOra(periodo, iIdx1, 0, numMedia)) & "," & _
                  ContaTutti_5_secondi(periodo, iIdx1, 0, numMedia) & "," & _
                  Replace(Str(StdDev(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                  Replace(Str(StdDev(periodo, iIdx1, 0, numMedia)), ",", ".") & "," & _
                  Replace(Str(DatoFlussoMassa(periodo, iIdx1, 1, numMedia)), ",", ".") & "," & _
                  .ParSQLStr("1" & Trim(Str$(gaConfigurazioneArchivio(iIdx1).STRUM.iddatabase))) & "," & stsImpianto & ")"
                  .ExecuteSQL (strSQL)
            End If
        End With
    End If
    
    Set rsDati = Nothing
    
    Exit Sub

GestErrore:
    Debug.Print Now & " ElaboraSalvaDatiSQL: " & Error(Err)
    Call WindasLog("BFdata ElaboraSalvaDatiSQL: " & Error(Err), 1)
    Set rsDati = Nothing

End Sub

Function ScegliDato(Ora, nn, numMedia, Status)

    Dim IndirizzoOrigine As Integer

    On Error GoTo GestErrore
    
    IndirizzoOrigine = nn
        
    'luca 15/09/2016
    'se parametro non presente imposto valore a -9999
    If IndirizzoOrigine < 0 Then
        ScegliDato = -9999
        Status = "ERR"
    Else
        'altrimenti se parametro presente -> scrivo valore e status calcolati (calcolati precedentemente)
        ScegliDato = MedieOra(Ora, IndirizzoOrigine, 0, numMedia)
        Status = StsMedieOra(Ora, IndirizzoOrigine, 0, numMedia)
    End If
    
    Exit Function

GestErrore:
    Call WindasLog("ScegliDato " + Error(Err), 1)
    Resume fine
fine:

End Function

Function StatusMonitor(statoimp() As Integer) As String


    Call QuickSort(statoimp, 8, True)
    
    If statoimp(0, 0) = 0 Then
        
        StatusMonitor = "ERR"
            
    Else
        'Alby Dicembre 2015
        Select Case statoimp(0, 1)
        'Select Case statoimp(0, 1)
    
            Case 1
                StatusMonitor = "ERR"
                
            Case 2
                StatusMonitor = "TZR"
                
            Case 3
                StatusMonitor = "TSP"
                
            Case 4
                StatusMonitor = "MAN"
                
            Case 5
                StatusMonitor = "OFF"
               
            Case 6
                StatusMonitor = "NVA"
               
            Case 7
                StatusMonitor = "NVL"
            
            Case 8
                StatusMonitor = "NVH"
                
            Case 9
                StatusMonitor = "TAR"
                
            Case Else
                StatusMonitor = "ERR"
    
        End Select
        
    End If
    


End Function

Sub WindasLog(evento As String, grave)

    Dim ll As Integer
    Dim rs As Object
    
    On Error GoTo GestErrore
    
    Debug.Print evento
    
    'Alby Febbraio 2012 log a video e su files giornalieri
    If Dir(App.Path & "\logBFdata", vbDirectory) = "" Then
        MkDir App.Path & "\logBFdata"
    End If
    
    ll = FreeFile
    Open App.Path + "\logBFdata\" + Format(Now, "dd-mmmyy") + ".txt" For Append As #ll
    Print #ll, Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + evento
    Close (ll)
    Form1.Label1.Caption = Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + evento
    
    Exit Sub
    
GestErrore:
    Debug.Print Error(Err)
    
End Sub
Function ProiezioneMediaGiornaliera(MediaInCorso)

    Dim Ore As Integer
    Dim OreMancanti As Integer
    Dim Proiezione As Double

    'Alby Dicembre 2015
    On Error GoTo GestErrore

    If MediaInCorso = -9999 Or UltimaMediaOraria = -9999 Then
        ProiezioneMediaGiornaliera = -9999
        Exit Function
    End If
    
    'luca marzo 2017
    If UCase(Tabella) = "WDS_ELAB" Then
        'luca marzo 2017 da verificare meglio, secondo me è sbagliata
        'Ore = Val(Format(Now, "hh")) - 1
        Ore = Val(Format(Now, "hh"))
        OreMancanti = 24 - Ore
        
        'luca marzo 2017
        If Ore > 0 Then
            Proiezione = ((MediaInCorso * Ore) + (UltimaMediaOraria * OreMancanti)) / 24
        Else
            Proiezione = -9999
        End If
    Else
        Ore = Int(DateDiff("n", DateSerial(year(Now), month(Now), day(Now)) + TimeSerial(0, 0, 0), Now) / 30)
        OreMancanti = 48 - Ore
        
        If Ore > 0 Then
            Proiezione = ((MediaInCorso * Ore) + (UltimaMediaOraria * OreMancanti)) / 48
        Else
            Proiezione = -9999
        End If
    End If
    ProiezioneMediaGiornaliera = Proiezione

    Exit Function

GestErrore:
    Call WindasLog("ProiezioneMediaGiornaliera ", 1)

End Function

Function ProiezioneMediaMensile(MediaInCorso)

    Dim Giorni As Integer
    Dim GiorniMancanti As Integer
    Dim Proiezione As Double

    'Alby Dicembre 2015
    On Error GoTo GestErrore

     If MediaInCorso = -9999 Or UltimaMediaGiorno = -9999 Then
        ProiezioneMediaMensile = -9999
        Exit Function
    End If

    'Alby Febbraio 2016
    'Giorni = Val(Format(Now, "hh")) - 1
    Giorni = Val(Format(Now, "dd"))
    GiorniMancanti = GiorniMese(Now) - Giorni
    
    Proiezione = ((MediaInCorso * Giorni) + (UltimaMediaGiorno * GiorniMancanti)) / GiorniMese(Now)
    ProiezioneMediaMensile = Proiezione

    Exit Function

GestErrore:
    Call WindasLog("ProiezioneMediaMensile ", 1)

End Function

'luca luglio 2017
Function ProiezioneMediaTrimestrale(MediaInCorso As Double, DataInizioTrimestre As Date, DataFineTrimestre As Date) As Double

    Dim GiorniTrascorsi As Integer
    Dim GiorniMancanti As Integer
    Dim Proiezione As Double
    Dim GiorniTotali As Integer

    On Error GoTo GestErrore

     If MediaInCorso = -9999 Or UltimaMediaGiorno = -9999 Then
        ProiezioneMediaTrimestrale = -9999
        Exit Function
    End If
    
    GiorniTotali = DateDiff("d", DataInizioTrimestre, DataFineTrimestre)
    GiorniTrascorsi = DateDiff("d", DataInizioTrimestre, Now)
    GiorniMancanti = GiorniTotali - GiorniTrascorsi
    
    Proiezione = ((MediaInCorso * GiorniTrascorsi) + (UltimaMediaGiorno * GiorniMancanti)) / GiorniTotali
    ProiezioneMediaTrimestrale = Proiezione

    Exit Function

GestErrore:
    Call WindasLog("ProiezioneMediaTrimestrale ", 1)

End Function
Function ProiezioneMedia48h()

    Dim Ore As Integer
    Dim OreMancanti As Integer
    Dim Proiezione As Double

    'Alby Dicembre 2015
    On Error GoTo GestErrore

    If MediaInCorso48h = -9999 Or UltimaMediaOraria = -9999 Then
        ProiezioneMedia48h = -9999
        Exit Function
    End If
    
    Ore = NrDatiInCorso48h
    OreMancanti = 48 - Ore
    
    Proiezione = ((MediaInCorso48h * NrDatiInCorso48h) + (UltimaMediaOraria * OreMancanti)) / 48
    ProiezioneMedia48h = Proiezione

    Exit Function

GestErrore:
    Call WindasLog("ProiezioneMedia48h ", 1)

End Function


Function CodParametro(Identificativo)

    Dim iIdx As Integer

    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    If Identificativo = -1 Then
        CodParametro = -1
        Exit Function
    End If
    
    CodParametro = -1
    For iIdx = 0 To gnNroParametriStrumenti
        If Identificativo = gaConfigurazioneArchivio(iIdx).STRUM.iddatabase Then
            CodParametro = iIdx
            Exit Function
        End If
    Next iIdx
    Call WindasLog("Attenzione parametro NON trovato", 1)
    
    Exit Function
    
GestErrore:
    Call WindasLog("CodificaStatus " + Error(Err), 1)


End Function

Sub ElaboraSalvaDati(Elabdate As Date)

    On Local Error GoTo GestErrore
    
    Call ElaboraSalvaDatiInizializzo
    
    'luca luglio 2017
    Call OPC.ChiudiOPC
    
    If gnNroParametriStrumenti >= 0 Then
    
        'Alby Luglio 2013 Enipower Bolgiano
        'NB matrice MedieOra(ora, parametro, x,                    tipo)
        '                                    0=tal quale misurato  0=mezzora
        '                                    1=elaborato           1=10 minuti
        '                                    2=tal quale stimato   2=ore
        '                                                          3=minuti
        'luca aprile 2017 routine in cui verifico la QAL2 da applicare dalla tabella wls_cfglog
        Call GestioneQAL2
        
        Call ElaboraSalvaDatiMedie

        'Alby Gennaio 2016 se necessario rielaborare
        Call ElaboraSalvaDatiStatoImpiantoPrevalente
        Call ElaboraSalvaDatiNormalizza
        
        'Ricalcolo H2O (NCX se uno dei due è invalido)
        Call ElaboraSalvaDatiCalcolaH2O
        
        Call ElaboraSalvaDatiFlussiMassa
        
        Call ElaboraSalvaDatiConcludo
        
    End If
    
    'luca luglio 2017
    Call OPC.ChiudiOPC
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraSalvaDati: " + Error(Err), 1)
    
End Sub



Function ElaborazioniDiLegge(ValoreTQ As Double, H2O As Double, O2 As Double, T As Double, P As Double, iIdx As Integer, Status As String)
 

    Const H2O_LEGGE = 100
    Const O2_LEGGE = 21
    Const TEMP_RIF = 273.15
    Const PRESS_RIF_PA = 1013.25    'mBar
    'Const PRESS_RIF_PA = 10332    'mm H2O
    Dim limiteO2 As Double
    Dim IngressoPortata As Integer
    Dim valoreElaborato As Double
    Dim O2riferimento As Double
    
    On Error GoTo GestErrore
    
    IngressoPortata = IngressoQFUMI
    
    'Alby Gennaio 2016
    O2riferimento = CDbl(Trim(Generiche(iO2RIF).Par))
    
    If ValoreTQ = -9999 Then ElaborazioniDiLegge = -9999: Exit Function
    ElaborazioniDiLegge = -9999
    valoreElaborato = ValoreTQ
    
    '************** applicazione retta di QAL2 *****************
    'daniele luglio 2013 bolgiano: riattivo qal2
    'luca marzo 2017 non rieseguo la QAL2 se attivata la QAL2 sul tal quale
    If Not gaConfigurazioneArchivio(iIdx).STRUM.QAL2suTQ Then
        If gaConfigurazioneArchivio(iIdx).STRUM.m <> 0 Then
            valoreElaborato = valoreElaborato * gaConfigurazioneArchivio(iIdx).STRUM.m + gaConfigurazioneArchivio(iIdx).STRUM.q
        End If
    End If
            
    'luca 05/10/2015 ripristino riporto al secco
    'Alby Novembre 2014
    '************* riporto della misura al secco ***************
    If InStr(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.Elaborazioni), left1(CONF_ELAB_SECCO, 1)) <> 0 Then
        'Alby Febbraio da controllare
        If H2O <> -9999 Then
            If iIdx = IngressoPortata Then
                valoreElaborato = valoreElaborato * (H2O_LEGGE - H2O) / H2O_LEGGE
            Else
                valoreElaborato = valoreElaborato * (H2O_LEGGE / (H2O_LEGGE - H2O))
            End If
        Else
            ElaborazioniDiLegge = -9999: Status = "NCU": Exit Function
        End If
    End If
    
    '************ Compensazione della Misura a Valore Noto di Ossigeno
    If InStr(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.Elaborazioni), left1(CONF_ELAB_COMP, 1)) <> 0 Then
        limiteO2 = 20
        If O2 > limiteO2 Then O2 = limiteO2
                                        
        Rem Controllo che Esista il Valore dell'Ossigeno in Acquisizione
        If O2 <> -9999 Then
            Rem Controllo che sia Maggiore di 0 e Non Uguale a 21 (Ossigeno
            Rem Presente in Aria per Legge) il Valore dell'Ossigeno in Acquisizione
            If O2 <= limiteO2 Then
                Rem riporto O2 al secco
                Rem alla funzione ElaboraDati DEVE entrare O2umido
                If iIdx = IngressoPortata Then
                    valoreElaborato = valoreElaborato * ((O2_LEGGE - O2) / (O2_LEGGE - O2riferimento))
                Else
                    valoreElaborato = valoreElaborato * ((O2_LEGGE - O2riferimento) / (O2_LEGGE - O2))
                End If
            Else
                ElaborazioniDiLegge = -9999: Status = "NCO": Exit Function
            End If
        Else
            ElaborazioniDiLegge = -9999: Status = "NCO": Exit Function
        End If
    End If
    
    '********** Normalizzazione della Misura per Temperatura
    If InStr(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.Elaborazioni), left1(CONF_ELAB_NORM, 1)) <> 0 Then
        
        'Alby Parma Febbraio 2013
        If T > -30 Then
        'If T > 0 Then
            If iIdx = IngressoPortata Then
                valoreElaborato = valoreElaborato * (TEMP_RIF / (T + TEMP_RIF))
            Else
                valoreElaborato = valoreElaborato * ((T + TEMP_RIF) / TEMP_RIF)
            End If
        Else
            ElaborazioniDiLegge = -9999: Status = "NCT": Exit Function
        End If
                    
        Rem Controllo che Esista il Valore della Pressione in Acquisizione
        If P > 0 Then
            If iIdx = IngressoPortata Then
                valoreElaborato = valoreElaborato * (P / PRESS_RIF_PA)
            Else
                valoreElaborato = valoreElaborato * (PRESS_RIF_PA / P)
            End If
        Else
            ElaborazioniDiLegge = -9999: Status = "NCP": Exit Function
        End If
    End If
    
    'daniele luglio 2013 bolgiano: attivo intervallo di confidenza
    If gaConfigurazioneArchivio(iIdx).STRUM.IntervalloConfidenza > 0 Then
        valoreElaborato = valoreElaborato - gaConfigurazioneArchivio(iIdx).STRUM.IntervalloConfidenza
    End If
    'daniele luglio 2013 bolgiano: attivo limite di rilevabilità
    If gaConfigurazioneArchivio(iIdx).STRUM.LimiteRilevabilita >= 0 Then 'Nicolò Settembre 2015 voglio anche lim rilevabilità a 0
        If valoreElaborato < gaConfigurazioneArchivio(iIdx).STRUM.LimiteRilevabilita Then
            valoreElaborato = gaConfigurazioneArchivio(iIdx).STRUM.LimiteRilevabilita
        End If
    End If
    
    ElaborazioniDiLegge = valoreElaborato
    
    Exit Function
    
GestErrore:
    Call WindasLog("ElaborazioniDiLegge " + Error(Err), 1)
    Resume fine

fine:

End Function

'Sub DatiSADSalvaSuFileXlinea(NomeFile, ElabDate, iOra, iSecondi)
'
'    Const ForReading = 1, ForWriting = 2, ForAppending = 8
'    Dim fso, f, MyFile, iIdx, DataFileARPA, OraFileARPA, nn, valore
'    Dim CodFile_4343(2), NomeSoftware_4343, NomeImp_4343, ValoreImpianto, stsImpianto
'    Dim NumLinea, sec_ndx, sec_ndx_str
'    Dim oraSAD
'    Dim minSAD
'    Dim iMinuti As Integer
'
'    On Error GoTo gestErrore
'
'    iMinuti = (iSecondi * 5) \ 60
'    sec_ndx = (iSecondi * 5) Mod 60
'    DataFileARPA = Format(ElabDate, "yyyymmdd")
'    OraFileARPA = Format(iOra, "00") & "." & Format(iMinuti, "00") & "." & Format(sec_ndx, "00")
'
'    '***** salvataggio files separati ****
'
'     Set fso = CreateObject("Scripting.FileSystemObject")
'     If (fso.FileExists(NomeFile)) Then
'
'         Set f = fso.OpenTextFile(NomeFile, ForAppending, True)
'
'         '***** Riga >=5 Data, Ora, valore istantaneo, stato della misura *****
'         f.Write DataFileARPA & Chr(9) & OraFileARPA
'
'         For iIdx = 0 To gnNroParametriStrumenti
'             DoEvents
'
'            If ValIst(0, iIdx) <= -8888 Then
'                valore = "---"
'            Else
'                valore = Trim(Replace(CStr(FormatNumber(ValIst(0, iIdx), 2, -2, -2, 0)), ",", "."))
'            End If
'            f.Write Chr(9) & valore & Chr(9) & Trim(CStr(Status(0, iIdx)))
'         Next iIdx
'
'         f.Writeline Chr(9)
'
'     Else
'
'         Set f = fso.OpenTextFile(NomeFile, ForWriting, True)
'
'         '***** Riga 1 Id. del software utilizzato dal gestore *****
'         f.Writeline Nome_Software_4343
'
'         '***** Riga 2 Codice Impianto assegnato da ARPA *****
'         f.Writeline Nome_Impianto_4343
'
'         '***** Riga 3 Codice Monitor *****
'         f.Write "#" & String(10, " ")
'         For iIdx = 0 To gnNroParametriStrumenti
'             'NumLinea = CInt(Right(Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro), 1))
'             'If nn = NumLinea Then
'                 'If Trim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX) = "" Or _
'                 '   UCase(Trim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX)) = "IMPIANTO" Then
'                 '    f.Write Chr(9) & Chr(9) & gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitor
'                 'Else
'                 '    f.Write Chr(9) & Chr(9) & gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitor
'                 '    f.Write Chr(9) & Chr(9) & gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX
'                 'End If
'             'End If
'             'Alby Luglio 2016
'             f.Write Chr(9) & Chr(9) & gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro
'         Next
'
'         f.Writeline Chr(9)
'
'         '***** Riga 4 Unità di misura dei Codici Monitor *****
'         f.Write "#" & String(10, " ")
'         For iIdx = 0 To gnNroParametriStrumenti
'             'NumLinea = CInt(Right(Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro), 1))
'             'If nn = NumLinea Then
'             '    If Trim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX) = "" Or UCase(Trim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX)) = "IMPIANTO" Then
'             '        f.Write Chr(9) & Chr(9) & Trim(gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisura)
'             '    Else
'                     f.Write Chr(9) & Chr(9) & Trim(gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisura)
'             '        f.Write Chr(9) & Chr(9) & Trim(gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisura)
'             '    End If
'             'End If
'         Next
'
'         f.Writeline Chr(9)
'
'         '***** Riga >=5 Data, Ora, valore istantaneo, stato della misura *****
'         f.Write DataFileARPA & Chr(9) & OraFileARPA
'         For iIdx = 0 To gnNroParametriStrumenti
'
'             'NumLinea = CInt(Right(Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro), 1))
'             'If nn = NumLinea Then
'                 'If UCase(Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro)) = "IMP_L" & CStr(Trim(NumLinea)) Then
'                 '    f.Write Chr(9) & ValIst(0, iIdx) & Chr(9) & "---"
'                 'Else
'                     If Trim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorAUX) = "" Then
'                         If ValIst(0, iIdx) <= -8888 Then
'                             valore = "---"
'                         Else
'                             valore = Trim(Replace(CStr(FormatNumber(ValIst(0, iIdx), 2, -2, -2, 0)), ",", "."))
'                         End If
'                         f.Write Chr(9) & valore & Chr(9) & Trim(CStr(Status(0, iIdx)))
'                     'Else
'                     '    If ValIst(0, iIdx) <= -8888 Then
'                     '        Valore = "---"
'                     '    Else
'                     '        Valore = Trim(Replace(CStr(FormatNumber(ValIst(0, iIdx), 2, -2, -2, 0)), ",", "."))
'                     '    End If
'                     '    f.Write Chr(9) & Valore & Chr(9) & Trim(CStr(Status(0, iIdx)))
'                     '    If ValIst(2, iIdx) <= -8888 Then
'                     '        Valore = "---"
'                     '    Else
'                     '        Valore = Trim(Replace(CStr(FormatNumber(ValIst(2, iIdx), 2, -2, -2, 0)), ",", "."))
'                     '    End If
'                     '    f.Write Chr(9) & Valore & Chr(9) & Trim(CStr(Status(2, iIdx)))
'                     'End If
'                 'End If
'
'             End If
'         Next
'
'         f.Writeline Chr(9)
'     End If
'
'     f.Close
'     Set f = Nothing
'     Set fso = Nothing
'
'    Exit Sub
'
'gestErrore:
'    Call WindasLog("BFdata DatiSADSalvaSuFileXlinea: " + Error(Err), 1)
'    Resume Next
'
'End Sub

Sub Elabora(Elabdate As Date)

    On Error GoTo GestErrore:
    
    'luca marzo 2017
    Call WindasLog("Iniziata elaborazione " & UCase(Tabella) & " --- " & CStr(Elabdate), 0)
    
    'Federica gennaio 2018 - Verifica connessione al Database
    Call CheckDBConnection
    
    If ConnessioneValida Then
    
        'luca 08/11/2016
         Call AzzeraMatrici
        
        'Federica gennaio 2018
        'Sempre con modo = 0 -> medie finite
        Call ElaboraCaricaDatiElementari(Elabdate, 0)
        
        If Tabella <> "WDS_10MINCO" And Tabella <> "WDS_AUTO" Then Call DatiSADSalva(Elabdate)
        
        Form1.Label1.Caption = "Elaboro dati " + Tabella
        Form1.Refresh
        Call ElaboraSalvaDati(Elabdate)
        
    Else
        Call WindasLog("Connessione al database non disponibile!", 0)
    End If
    
    Call WindasLog("Terminata elaborazione " & UCase(Tabella) & " --- " & CStr(Elabdate), 0)
    
    Exit Sub

GestErrore:
    Debug.Print Now & " Elabora: " & Error(Err)
    Call WindasLog("BFdata Elabora: " + Error(Err), 1)
    Resume fine:
fine:

End Sub


Function RicavoRuolo() As Integer

    Dim nFile As Integer
    Dim riga As String
    
    On Error GoTo GestErrore
    
    'Alby Luglio 2016
    'se presente file ini
    If Dir(App.Path & "\Ruolo.ini") <> "" Then
        nFile = FreeFile
        Open App.Path & "\Ruolo.ini" For Input As #nFile
        Line Input #nFile, riga
        'Alby Gennaio 2016
        RicavoRuolo = CInt(Trim(riga))
        Close #nFile
    Else
        RicavoRuolo = 0
    End If

    Exit Function
    
GestErrore:
    Call WindasLog("RicavoRuolo " + Error(Err), 1)

End Function
Function IsClient() As Boolean

    Dim nFile As Integer
    Dim riga As String
    
    On Error GoTo GestErrore
    
    'Alby Luglio 2016
    'se presente file ini
    If Dir(App.Path & "\Client.ini") <> "" Then
        nFile = FreeFile
        Open App.Path & "\Client.ini" For Input As #nFile
        Line Input #nFile, riga
        'Alby Gennaio 2016
        IsClient = CBool(Trim(riga))
        Close #nFile
    Else
        IsClient = False
    End If

    Exit Function
    
GestErrore:
    Call WindasLog("IsClient " + Error(Err), 1)

End Function
Private Sub AzzeraMatrici()
    
    Dim hh As Integer
    Dim I2 As Integer
    Dim iIdx As Integer
    Dim ss As Integer
    
    'Alby Ottobre 2016
    On Error GoTo GestErrore
    
    '***** reset variabili *****
    For hh = 0 To 23
        For iIdx = 0 To gnNroParametriStrumenti
            For ss = 0 To 720
                Valore_5_Secondi(hh, iIdx, ss) = -9999
                Valore_5_Secondi_N(hh, iIdx, ss) = -9999
                Valore_5_Secondi_S(hh, iIdx, ss) = -9999
                Valore_5_Secondi_SN(hh, iIdx, ss) = -9999
                ContaTuttiSecondiMediaOra(hh, iIdx, ss) = 0
                Status_5_Secondi(hh, iIdx, ss) = "-9999"
                Status_5_Secondi_N(hh, iIdx, ss) = "-9999"
                Status_5_Secondi_S(hh, iIdx, ss) = "-9999"
                Status_5_Secondi_SN(hh, iIdx, ss) = "-9999"
            Next ss
        Next iIdx
        
        'Alby Giugno 2016 azzeramento stato impianto
        For I2 = 0 To 65
            statoimp(hh, I2) = 0
            PercRegime(hh, I2) = 0
            PercMinTec(hh, I2) = 0
            PercSpegnimento(hh, I2) = 0
            PercManutenzione(hh, I2) = 0
            PercFermo(hh, I2) = 0
            PercGuasto(hh, I2) = 0
            PercAnomalo(hh, I2) = 0
            PercPolveri(hh, I2) = 0
            PercAltro(hh, I2) = 0
        Next I2
    Next hh

    Exit Sub

GestErrore:
    Call WindasLog("BFdata AzzeraMatrici: " + Error(Err), 1)

End Sub

Function left1(stringa As Variant, Numero As Variant)
    
    left1 = Mid(stringa, 1, Numero)

End Function

Function right1(stringa As Variant, Numero As Variant)

    right1 = Mid(stringa, Len(stringa) - Numero + 1)

End Function

'Sub OLD_GetConnectionParams(Connessione)
'    Dim param As Object
'  Dim Crypt As Object
'  Dim ErrCount As Integer
'  Dim AccessObj As Object
'  Dim ErrPos As Integer
'  Dim AppReadFromIni As Boolean
'
'  AppReadFromIni = True
'  If (Dir$(App.Path & "\bflab7.mdb") <> "") Then
'    Set AccessObj = CreateObject("AttimoFwk.CData")
'    With AccessObj
'
'RetryForProviderErr:
'      .Err_Activate = True
'      .SetMessages False
'      Select Case ErrCount
'        Case 0
'          .SetDBType .Conn_Jet
'          .SetDatabase App.Path & "\bflab7.mdb", "admin", ""
'        Case Else
'          AppReadFromIni = True
'      End Select
'      AppReadFromIni = False
'      If (Not AppReadFromIni) Then
'        AppReadFromIni = True
'
'        'Alby Ottobre 2016 attenzione per ridondanza configurare 2 e solo 2 connessioni
'        If Connessione = 1 Then
'            'configurare come default con BFconfigCN sempre il DB locale
'            .SelectionFast ("SELECT * FROM Connections WHERE DEFAULT = -1")
'        Else
'            'configurare senza default l'altro DB (partner ridondanza)
'            .SelectionFast ("SELECT * FROM Connections WHERE DEFAULT <> -1")
'        End If
'
'        If (Not .IsEOF) Then
'          AppServer = .GetValue("SERVER")
'          AppDatabase = .GetValue("DATABASE")
'          AppDBType = .GetValue("DBTYPE")
'          AppDbVersion = .GetValue("CNMODE")
'          AppDBUser = .GetValue("USER")
'          AppDBPwd = .GetValue("PASSWORD")
'
'          Set Crypt = CreateObject("AttimoFwk.CCrypt")
'          AppDBUser = Crypt.Decrypt(AppDBUser)
'          AppDBPwd = Crypt.Decrypt(AppDBPwd)
'          Set Crypt = Nothing
'          AppReadFromIni = False
'
'        End If
'      End If
'    End With
'    Set AccessObj = Nothing
'
'  End If
'
'  If (AppReadFromIni) Then
'    Set param = CreateObject("AttimoFwk.CParam")
'
'    With param
'      .ParamFile = App.Path & "\bflab7.ini"
'      AppServer = .GetStringIni("SERVER")
'      AppDatabase = .GetStringIni("DATABASE")
'      AppDBType = .GetStringIni("DBTYPE")
'
'      AppDBUser = .GetStringIni("USER")
'      AppDBPwd = .GetStringIni("PASSWORD")
'      Set Crypt = CreateObject("AttimoFwk.CCrypt")
'      AppDBUser = Crypt.Decrypt(AppDBUser)
'      AppDBPwd = Crypt.Decrypt(AppDBPwd)
'
'      Set Crypt = Nothing
'    End With
'
'    Set param = Nothing
'  End If
'
'End Sub

Function SuperTrim(ByVal StringToTrim As String) As String
    
    '** La funzione sostituisce tutti i Chr$(0) di una stringa con Spazi **
    '** e quindi restituisce tale stringa "Trimmata"                     **
    
    StringToTrim = SostituisciCarattere(StringToTrim, Chr$(0), Chr$(32))
    SuperTrim = Trim$(StringToTrim)

End Function
Function SostituisciCarattere(ByVal stringa As String, ByVal StrRicerca As String, ByVal StrSostituire As String) As String

Dim nPointer As Integer
Dim nNextStart As Integer
Dim sBuffer As String
    
'*** Sostituisce tutte le ricorrenze di StrRicerca con StrSostituire
'*** Parametri:
'***    Stringa è la stringa su cui operare
'***    StrRicerca è la stringa da ricercare
'***    StrSostituire è la stringa da inserire al posto di StrRicerca
    
    If StrRicerca = "" Then
        SostituisciCarattere = stringa
        Exit Function
    End If
    nNextStart = 1
    Do While InStr(nNextStart, stringa, StrRicerca) > 0
        nPointer = InStr(nNextStart, stringa, StrRicerca)
        If nPointer Then
            sBuffer = right1(stringa, Len(stringa) - (nPointer + Len(StrRicerca) - 1))
            stringa = left1(stringa, InStr(nNextStart, stringa, StrRicerca) - 1)
            stringa = stringa + StrSostituire + sBuffer
            nNextStart = nPointer + Len(StrSostituire)
        End If
    Loop
    SostituisciCarattere = stringa

End Function

Public Sub Ritardo(ByVal fSecondi As Single)

    Dim dAttesa As Double
    
    Rem ***** Ciclo di Ritardo *****
    dAttesa = Timer
    Do
        DoEvents
    Loop Until Mezzanotte(dAttesa) > fSecondi

End Sub
Public Function Mezzanotte(ByVal dAttesa As Double) As Double
    
    Dim dLetto As Double
    
    dLetto = Timer
    Mezzanotte = dLetto - dAttesa
    
    If CLng(Mezzanotte) < 0 Then
        Mezzanotte = 86400 - dAttesa + dLetto
    End If

End Function


Sub ElaboraSalvaDatiMedieNF(Ore, parametro, Elabdate)
    
    'Alby Dicembre 2015
    Dim rsDati As Object
    Dim rsDati2 As Object
    Dim rs As Object
    Dim strSQL As String
    Dim UltimaData As String
    Dim UltimaOra As String
    Dim NrOre As Integer
    Dim Sommatoria As Double
    Dim Contatore As Integer
    Dim Media As Double
    Dim Status As String
    Dim NrOra As Integer
    
    On Error GoTo GestErrore
    
    NewDataObj rsDati
    NewDataObj rsDati2
    
    'setto ultima data ad inizio anno in corso
    UltimaData = Format(Elabdate, "yyyy") + "0101"
    UltimaOra = "00.00"
    UltimaMedia48h = -9999
    IDUltima48H = 0
    IDCostruzione48H = 0
    
    
    'If parametro = 20 Then Stop
    'determino ultima media registrata
    'daniele dicembre 2013 tirreno power: a inizio anno le medie delle 12/48 ore devono ripartire da zero (richiesta esplicita del cliente)
    'strSQL = "SELECT dt_date,dt_hour FROM wds_" + Trim(Str(Ore)) + "h WHERE dt_stationcode='" + StationCode + "' and dt_measurecod='" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "' " + "AND DT_date >= '" + UltimaData + "' AND DT_Nrtot = " + RTrim(Str(Ore)) + " ORDER BY dt_date, dt_hour"
    strSQL = "SELECT * FROM wds_" + Trim(Str(Ore)) + "h WHERE dt_stationcode='" + StationCode + "' and dt_measurecod='" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "' AND DT_Nrtot = " + RTrim(Str(Ore)) + " ORDER BY dt_date, dt_hour"
    
    With rsDati
        If (.SelectionFast(strSQL)) Then
            
            Do While Not .iseof
                .MoveNext
            Loop
            
            'Alby Dicembre 2015
            UltimaData = .GetValue("dt_date")
            UltimaOra = .GetValue("dt_hour")
            UltimaMedia48h = .GetValue("dt_value")
            
            'luca 06/09/2016 calcolo ID ultima media 48H
            '*************** ID
            Dim OreTotaliMarcia As Double
            Dim OreTotaliMarciaValide As Double
            
            OreTotaliMarcia = .GetValue("DT_Nrtot")
            OreTotaliMarciaValide = .GetValue("DT_Nr")
            
            If OreTotaliMarcia > 0 Then
                IDUltima48H = OreTotaliMarciaValide / OreTotaliMarcia * 100
                If IDUltima48H > 100 Then IDUltima48H = 100
                'luca 16/09/2016 gestisco status Ultima 48H
                If IDUltima48H >= 70 Then
                    StatusUltima48H = "VAL"
                Else
                    StatusUltima48H = "ERR"
                End If
            Else
                IDUltima48H = 0
                'luca 16/09/2016 gestisco status Ultima 48H
                StatusUltima48H = "ERR"
            End If
            
        End If
    End With
    
    'Alby Dicembre 2015 cancella ultima media in costruzione
    'luca 08/11/2016 non lo faccio se sono il client
    If Not Client Then
        strSQL = "DELETE FROM wds_" + Trim(Str(Ore)) + "h WHERE DT_NrTot < " + Trim(Str(Ore)) + " and dt_measurecod='" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "'"
        rsDati2.ExecuteSQL (strSQL)
    End If
    
    'scorro i dati dall'ultima media registrata fino ad arrivare alle ore richieste
    strSQL = "SELECT * FROM wds_elab WHERE dt_stationcode='" + StationCode + "' and dt_measurecod='" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "' AND dt_date>='" + UltimaData + "' ORDER BY dt_date, dt_hour"
    With rsDati
        If (.SelectionFast(strSQL)) Then
            Do While Not .iseof
                'se ciclando sono arrivato oltre la data da elaborare esco dal ciclo operazioni ultimate
                If .GetValue("dt_date") > Format(Elabdate, "yyyymmdd") Then Exit Do
                
                'se l'ora elaborata è maggiore dell'ultima ora elaborata inizio le operazioni
                'Alby Gennaio 2016 se la data del record che scorre è > ultimaData sono passato al giorno successivo quindi posso azzerare UltimaOra
                If .GetValue("dt_date") > UltimaData Then UltimaOra = ""
                If .GetValue("dt_hour") > UltimaOra Or UltimaOra = "" Then
                    UltimaOra = ""
                    If .GetValue("dt_custom1") = "30" Then
                        NrOra = NrOra + 1
                        'Alby Dicembre 2013 da valutare se fare media anche con dati tal quale
                        If (.GetValue("dt_validflag") = "VAL" Or .GetValue("dt_validflag") = "AUX" Or .GetValue("dt_validflag") = "UTN") And .GetValue("dt_value") <> -9999 Then
                            
                            'Alby Dicembre 2015
                            'If parametro = 21 Then Stop

                            Sommatoria = Sommatoria + .GetValue("dt_value")
                            Contatore = Contatore + 1
                        End If
                        
                        If NrOra = Ore Then
                            Status = "ERR"
                            If Contatore > 0 Then
                                'luca 16/09/2016 calcolo la media anche con ID < 70% e gestisco la validità con lo status
                                Media = Sommatoria / Contatore
                                IDUltima48H = Contatore / NrOra * 100
                                If IDUltima48H >= 70 Then
                                    Status = "VAL"
                                End If
                            'luca 16/09/2016 se contatore pari a 0 media a -9999
                            Else
                                Media = -9999
                                IDUltima48H = 0
                            End If
                            
                            UltimaMedia48h = Media
                            StatusUltima48H = Status
                            
                            'luca 08/11/2016 non lo faccio se sono il client
                            If Not Client Then
                                strSQL = "INSERT INTO wds_" + Trim(Str(Ore)) + "h (dt_stationcode,dt_measurecod,dt_date,dt_hour,dt_value,dt_validflag, dt_nr, dt_nrtot)"
                                strSQL = strSQL + " VALUES ('" + StationCode + "','" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "','" + .GetValue("dt_date") + "','"
                                strSQL = strSQL + .GetValue("dt_hour") + "'," + Trim(Str(Media)) + ",'" + Status + "'," + Trim(Str(Contatore)) + "," + Str(NrOra) + ")"
                                rsDati2.ExecuteSQL strSQL
                            End If
                            
                            'Nicolò Tirreno Power resetto variabili dopo aver registrato una media completa
                            NrOra = 0
                            Sommatoria = 0
                            Contatore = 0
                            Media = -9999
                            Status = "Err"
                        End If
                    End If
                End If
                .MoveNext
            Loop
            
            'scrivo ultima media in costruzione
            MediaInCorso48h = -9999
            'luca 16/09/2016
            IDCostruzione48H = 0
            StatusCostruzione48H = "ERR"
            
            If NrOra > 0 Then
                
                Status = "ERR"
                If Contatore > 0 Then
                    Media = Sommatoria / Contatore
                    'Alby Dicembre 2015
                    MediaInCorso48h = Media
                    NrDatiInCorso48h = Contatore
                    'luca 06/09/2016 ID media 48H in costruzione
                    '****************** ID
                    IDCostruzione48H = Contatore / NrOra * 100
                    If IDCostruzione48H > 100 Then IDCostruzione48H = 100
                    'luca 16/09/2016 gestisco validità media 48H in costruzione
                    If IDCostruzione48H >= 70 Then
                        Status = "VAL"
                    End If
                End If
                
                StatusCostruzione48H = Status
                
                If Not Client Then
                    strSQL = "INSERT INTO wds_" + Trim(Str(Ore)) + "h (dt_stationcode,dt_measurecod,dt_date,dt_hour,dt_value,dt_validflag, dt_nr, dt_nrtot)"
                    strSQL = strSQL + " VALUES ('" + StationCode + "','" + Trim(gaConfigurazioneArchivio(parametro).STRUM.NomeParametro) + "','" + Format(Elabdate, "yyyymmdd") + "',"
                    strSQL = strSQL + Format(Elabdate, "hh") + ".00," + Trim(Str(Media)) + ",'" + Status + "'," + Trim(Str(Contatore)) + "," + Str(NrOra) + ")"
                    rsDati2.ExecuteSQL strSQL
                End If
            End If
            
        End If
        
    End With
    
    Set rsDati = Nothing
    Set rsDati2 = Nothing
    
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraSalvaDatiMedieNF " + Error(Err), 1)
    Resume Next

End Sub



