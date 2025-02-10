Attribute VB_Name = "VariabiliPubbliche"
Option Explicit

'**********************************************
'*      DICHIARAZIONE VARIABILI PUBBLICHE     *
'**********************************************

'*** GENERICHE ***
Public strSQL As String
Public sec_ndx
Public Refresh As Boolean
Public PrimaVolta As Boolean

'*** INFO GENERALI ***
Public NumeroLinea As Integer   'luca 21/07/2016
Public NomeLinea As String
Public Client As Boolean
Public gsDirLavoro As String
Public gsClienteDi As String    'Nicolò Agosto 2016 diventa global
Public gsImpianto As String
Public PCFunction As String
Public IsMaster As Boolean  'Federica luglio 2018

'*** 4343 ***
Public Nome_Software_4343 As String
Public Nome_Impianto_4343 As String
Public Nome_File_4343 As String
Public Const i4343_File = 100
Public Const i4343_Impianto = 101
Public Const i4343_SW = 102

'*** CONFIGURAZIONI ***
Public OreSemiore As Integer
Public GestioneSonoroAbilitata As Boolean
Public GestioneSMECloudAbilitata As Boolean
Public strValid As String
Public ElencoMisureStimate As String    'Federica ottobre 2017
Public AbilitaWatchdogPLC As Boolean    'Federica gennaio 2018
Public AbilitaHotBackup As Boolean      'Federica giugno 2018

'*** DATI ELEMENTARI ***
Public PathDAT As String

'*** DATI PER CONNESSIONE A DATABASE ***
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
Public iConnDBDefault As Integer

'*** PLC ***
Public DaLeggereRegistriPLC As Boolean  'Alby Luglio 2016

'*** ASSEGNAZIONE PARAMETRI ***
Public IngressoNO As Integer
Public IngressoNO2 As Integer   'Federica gennaio 2018 - Per calcolo NOX
Public IngressoNOX As Integer
Public IngressoO2 As Integer
Public IngressoO2Umido As Integer   'Per il calcolo dell'H2O
Public IngressoTemp As Integer
Public IngressoPress As Integer
Public IngressoH2O As Integer
Public IngressoPortata As Integer
Public IngressoStatoImpianto As Integer
Public IngressoVelocita As Integer  'Per il calcolo della portata
'luca marzo 2018
Public IngressoNH3 As Integer
Public IngressoNOXNH3 As Integer
Public IngressoDeltaP As Integer

'*** GESTIONE PARAMETRI ***
Public gnNroParametriStrumenti As Integer 'Nicolò Agosto 2016 diventa Public
Public nroDigitali As Integer
Public Status(3, 72)
Public ValIst(3, 72) '0=ingegnerizzato; 1=normalizzato; 3=grezzo
Public manValoreDigitale(999, 999)  'Alby Giugno 2014
Public Valore_DI(300)
Public ValPerc(3, 72)   'Federica ottobre 2017

Public MediaOraInCorso(2, 72)
Public StatusMediaOraInCorso(2, 72)
Public ID_MediaOraInCorso(1, 72) As Double 'luca 06/09/2016 ID Ora in corso normalizzata

'*** ANALOGICHE ***
Public Enum TipiAcquisizione
        PLC = 0
        SERIALE = 2
        CALCOLATO = 3
End Enum

' *Generali
'Federica gennaio 2018 - Tipizzo
Public Type Parameter
    CodiceParametro As String
    NomeParametro As String
    DescrParametro As String
    UnitaMisura As String
    NroDecimali As Integer
    ISE As Double
    FSE As Double
    ISI As Double
    FSI As Double
    FSI2 As Double
    SogliaAttenzione As Double
    SogliaAllarme As Double
    LimiteInferiore As Double
    LimiteSuperiore As Double
    LimiteInferioreOrario As Double
    LimiteSuperioreOrario As Double
    Acquisizione As Boolean
    TipoAcquisizione As TipiAcquisizione
    OpzioniAcquisizione As Variant
    TipoStrumento As Variant    'Non viene usato
    NroMorsetto As Integer
    SogliaValidazione As Variant    'Non viene usato
    idDatabase As Integer
    Elaborazioni As String
    MaxIncremento As Variant    'Non viene usato
    MinEscursione As Variant    'Non viene usato
    MaxEscursione As Variant    'Non viene usato
    LimConcMediaSemiorariaA As Double
    LimConcMediaOraria As Double
    LimConcMediaGiornaliera As Double
    LimConcMedia48H As Variant  'usata in BFData (ex LimiteMediaOraria_1)
    LimConcMediaMensile As Double
    Invalida As String
    LimiteRilevabilita As Double
    m As Double
    q As Double
    Range As Double
    IntervalloConfidenza As Double
    DataQAL2 As String
    ZeroSams As Double
    SpanSams As Double
    NomeTagDizionario As String 'luca 21/07/2016 (per versione 1)
    Precisione As Variant   'Non viene usato
    SogliaIstMin As Variant 'Non viene usato
    SogliaIstMax As Variant 'Non viene usato
    LimConcMediaAnnuale As Double
    ZeroTeorico As Double
    SpanTeorico As Double
    AttivaControlloConfigurazioneSoglie As Boolean  'luca 25/07/2016
    AttivaControlloConfigurazioneQAL2QAL3 As Boolean    'luca 25/07/2016
    AttivaControlloConfigurazioneValoreStimato As Boolean   'luca 25/07/2016
    QAL2suTQ As Boolean 'luca aprile 2017
    ErroreZero As Variant   'Non viene usato
    ErroreSpan As Variant   'Non viene usato
    CodiceMonitorIst_TQ As String
    CodiceMonitorMed_TQ As String
    CodiceMonitorMed_EL As String
    OrdineParametriADIADM As Integer
    SogliaAttenzioneGiornaliera As Double
    SogliaAllarmeGiornaliera As Double
    LimConcMediaTrimestrale As Double
    FattoreConversione As Double 'Alby Agosto 2017
    SogliaAttenzioneMensile As Double
    SogliaAllarmeMensile As Double
    IndiceDigitale2CampoScala As Integer    'Federica febbraio 2018
    IndiceTagCollegataQAL2 As Integer   'Federica giugno 2018
End Type
Public ParametriStrumenti(100) As Parameter

'*** DIGITALI ***
Public CodiceParametro_DI(300)
Public NomeParametro_DI(300)
Public Famiglia_DI(300)
Public NroMorsetto_DI(300, 1) As Integer    'Alby Giugno 2014
Public Contatti_DI(300)
Public StatoLogico_DI(300)
Public Colore0_DI(300)
Public Testo0_DI(300)
Public Colore1_DI(300)
Public Testo1_DI(300)
Public Priorita_DI(300)
Public Sonoro_DI(300) As Integer    'luca luglio 2017
Public IndiceDO(1000, 1000) 'Federica dicembre 2017

'Federica gennaio 2018
Type LineaGeneriche
    Par As Variant
    Testo As String
    Descrizione As String
End Type
Global Generiche(10000) As LineaGeneriche
'Indici parametri particolari
Public Const iO2RIF = 0
Public Const iStatoImpianto = 20
Public Const iPortata = 21
Public Const iTemperatura = 22
Public Const iPressione = 23
Public Const iH2O = 24
Public Const iO2 = 25
Public Const iO2Umido = 26
Public Const iNO = 27
Public Const iNOX = 28
Public Const iVelocita = 29
Public Const iNO2 = 30
Public Const iDeltaP = 31
Public Const iNrByte = 40
Public Const iMisureStimate = 45
Public Const iMisureQAL3 = 46
Public Const iSMECloudParametri = 47
Public Const iSMECloudAllarmi = 48
Public Const iIP_PLC = 52
Public Const iOreSemiore = 53
Public Const i10Minuti = 54
Public Const iFileSonoro = 55
Public Const iDivisorePerFlussi = 56
Public Const iWDPLC = 59
Public Const iRaggioCamino = 60
Public Const iMisureSimulate = 98
'luca marzo 2018
Public Const iNH3 = 31
Public Const iNOXNH3 = 32
'luca maggio 2018
Public Const iElencoAllarmiSuperoPLC = 61
Public Const Kportata = 62
Public Const AreaCamino = 71
'Federica giugno 2018
Public Const iHotBackup = 81

'*** BFDATA ***
Public EseguiMedie As Boolean
Public EseguitMedie10MinutiCO As Boolean   'luca aprile 2017

'*** RECUPERO DATI ***
Public PathBFImport As String
Public PathFileImportResult As String
Public AdamPath As String

