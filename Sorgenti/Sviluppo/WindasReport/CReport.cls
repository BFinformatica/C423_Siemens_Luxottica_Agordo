VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Property **
Public m_DbType As String
Public m_DbDatabase As String
Public m_DbUser As String
Public m_DbPassword As String
Public m_DbServer As String
Public m_DbVersion As String
Public m_StartDate As Date
Public m_EndDate As Date
Public m_Stations As String
Public m_Param As String
Public m_Table As String
Public m_ModelFileName As String
Public vet As Variant
Public decimali As Integer
        
'*********************************************
'*       sezione dimensionamento report      *
'*********************************************
'Federica luglio 2017 - gestione posizionamento su modello >>>
Public offset_riga As Integer
Public offset_colonna As Integer
Const POS_COL_NOTE = 0
Const POS_COL_VALORE = 1
Const POS_COL_ID = 2
Const CELLA_NOME_IMPIANTO = "B2"
Const CELLA_DATA = "AK2"
Const CELLA_TIPO_REPORT = "B4"
Const RIGA_PARAMETRI = 5
Const RIGA_INTESTAZIONI = 6
Const COLONNA_INTESTAZIONI = 1
Const COLONNA_PARAMETRI = COLONNA_INTESTAZIONI + 1
Const COLONNE_PER_PARAMETRO = 3       'definisce il numero di colonne utilizzate prt ogni parametro
Public PARAMETRI_PER_PAGINA As Integer   'definisce per quanti parametri per pagina è formattato il modello
Public COLONNA_IMPIANTO As Integer      'definisce la colonna, nel modello, dove riportare lo stato impianto o le ore di norm. funz.
Public NR_RIGHE As Integer      'Nr righe di dettaglio G = 24/48; M = 31; A = 12
Public RIGA_LIMITI_SINGOLI As Integer
Public RIGA_MINIMO As Integer
Public RIGA_MASSIMO As Integer
Public RIGA_TOTALE_PERIODO As Integer
Public RIGA_LIMITI_PERIODO As Integer
Public RIGA_CONTROLLO_97 As Integer
Public RIGA_LIMITI_97 As Integer
Public RIGA_LEGENDA_NOTE As Integer
Public COLONNA_LEGENDA_NOTE As Integer
Public CELLA_ELAB_LEGGE As String
Public CELLA_O2_RIF As String
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
'Nicola 13/01/2015 Variabili per copiare la riga sotto il nome del parametro
'Variabili QAL2
Public xStartRange, yStartRange, xEndRange, yEndRange, DestTopLeft, yDestTopLeft

Public numLinea
    
'Nicola 05/03/2015 Variabili Publi QAL3
'***************** Estrazione dati generali AMS *****************
Public gt_code
Public par

'****************** Estrazione scarto tipo AMS ******************
Public ZERO_Sams
Public SPAN_Sams

'Variabili Public per funzione IndiceOra
Public hh
Public mm
'Variabili Public per funzione CoordToString
Public result
Public charOfset
    
'Federica luglio 2017
Dim xMin As Double
Dim xMax As Double
    
Dim TotOreNormFunz As Integer

'luca 04/06/2015 variabili per la gestione orario/semiorario
Dim NumMedieOrario As Integer
Dim NumMedieSemiorario As Integer
Dim MinimoOreGiorno As Integer
Dim MinimoSemioreGiorno As Integer
Dim MinimoOreMese As Integer
Dim MinimoSemioreMese As Integer
Dim MinimoOreTrimestre As Integer 'Federica luglio 2017

Dim objEngine As Object
    
'*********************************************

'*********************************************
'*       vettori legati ai parametri         *
'*********************************************
Dim CodPar(99, 5) As String
Dim LimiteSemiora100(99) As Double
Dim LimiteSemiora97(99) As Double           'alias limite semiorario colonna B
Dim LimiteOra(99) As Double                 'Usato anche come limite semiorario colonna A se il report è semiorario
Dim LimiteGiorno(99) As Double
Dim LimiteMese(99) As Double
Dim LimiteTrimestre(99) As Double
Dim LimiteAnno(99) As Double
Dim Limite48H(99) As Double
'luca 19/01/2016
Dim LimiteOraFlussi(99) As Double
Dim LimiteMeseFlussi(99) As Double
Dim LimiteGiornoFlussi(99) As Double
Dim LimiteAnnoFlussi(99) As Double
Dim LimiteTrimestreFlussi(99) As Double
    
'Nicola
Public CheckDate As Date
Public String_StartDate_DB As String
Public String_EndDate_DB As String
    
Public NomeImpianto As String
Public rsExtractor As Object
Public rsValMin As Object
Public rsNfunz As Object
    
Public strValid As String
    
Public strSQL As String
Public CurRow As Integer
Public CurCol As Integer
Public SelStation As String
Public SelParam As String
Public SelTable As String
Public NFunzPar As String
Dim NFunzValue(48) As Double
Public PortataPar As String
'TODO: Da utilizzare
Public FattorePortata As Double

'TODO: Da verificare
Public maxNumeroAcquisizioni As Integer
    
'** Indici per cicli **
Public z As Integer
Public ndx As Integer
Public i As Integer
Public MaxPar As Integer 'Inizializzato con "Ubound(array parametri)"

Public OreNormFunz As Integer
    
'TODO: Spostare nella routine dove viene usata
Public RigaAlarm As Integer

Dim ListaParametri() As String
Public O2_rif As String
Public O2_rif_s As String
Public Abilita10minCO As Boolean 'Federica settembre 2017
    
'TODO: Da ragionare
Public riga As Integer
    
'Federica giugno 2017 - Se TRUE stampa le annotazioni del Piemonte
Const IsPiemonte = False

Public media_ As Double
Public note_ As String
Public id_ As Double

Option Explicit
'Federica settembre 2017
Private Function Controllo97(ByVal indice As Integer, ByVal anno As Integer) As Double

    Dim strSQL As String
    Dim NumValidi As Integer
    Dim NumMinori As Integer
    
    On Error GoTo GestErrore
    
    'Conteggio semiore in marcia per l'anno
    strSQL = "SELECT COUNT(*) AS Tot FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "AND LEFT(DT_DATE, 4) = '" & CStr(anno) & "' AND DT_CUSTOM1 = 30"
    If rsExtractor.selectionfast(strSQL) Then
        If Not rsExtractor.IsEOF Then NumValidi = rsExtractor.GetValue("Tot")
    End If
    
    If LimiteSemiora97(indice) <> 0 Then
        'Conteggio semiore valide del parametro
        strSQL = "SELECT COUNT(*) AS Tot FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_MEASURECOD = '" & CodPar(indice, 0) & "' " & _
                 "AND LEFT(DT_DATE, 4) = '" & CStr(anno) & "' AND DT_VALIDFLAG IN ('VAL', 'AUX') AND DT_VALUE <= " & LimiteSemiora97(indice)
        If rsExtractor.selectionfast(strSQL) Then
            If Not rsExtractor.IsEOF Then NumMinori = rsExtractor.GetValue("Tot")
        End If
    Else
        NumMinori = 0
    End If
    
    If NumValidi > 0 Then
        Controllo97 = NumMinori / NumValidi * 100
    Else
        Controllo97 = 0
    End If
    
    Exit Function
GestErrore:
    Call windasLog("Controllo97: " & Error(Err()))
    Controllo97 = 0

End Function

'Federica settembre 2017
Private Sub Elabora10MinutiCO(ByVal parametro_CO As String)

    Dim CO As String
    Dim strSQL As String
    Dim NumTot10MinCO As Integer
    Dim NumValidi10MinCO As Integer
    Dim NumSupero10MinCO As Integer

    On Error GoTo GestErrore
    CO = parametro_CO
    
    'Conteggio 10 minuti TOTALI del giorno
    strSQL = "SELECT COUNT(*) AS Tot FROM Wds_10minCO WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_MEASURECOD = '" & CO & "' " & _
             "AND DT_DATE = '" & String_StartDate_DB & "' AND DT_CUSTOM1 = 30 " & _
             "GROUP BY DT_MEASURECOD"
    If rsExtractor.selectionfast(strSQL) Then
        If Not rsExtractor.IsEOF Then NumTot10MinCO = rsExtractor.GetValue("Tot")
    End If
    
    'Conteggio 10 minuti VALIDI del giorno
    strSQL = "SELECT COUNT(*) AS Tot FROM Wds_10minCO WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_MEASURECOD = '" & CO & "' " & _
             "AND DT_DATE = '" & String_StartDate_DB & "' AND DT_CUSTOM1 = 30 AND DT_VALIDFLAG = 'VAL' " & _
             "GROUP BY DT_MEASURECOD"
    If rsExtractor.selectionfast(strSQL) Then
        If Not rsExtractor.IsEOF Then NumValidi10MinCO = rsExtractor.GetValue("Tot")
    End If
    
    'Conteggio 10 minuti SUPERI del giorno
    strSQL = "SELECT COUNT(*) AS Tot FROM Wds_10minCO WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_MEASURECOD = '" & CO & "' " & _
             "AND DT_DATE = '" & String_StartDate_DB & "' AND DT_CUSTOM1 = 30 AND DT_VALIDFLAG = 'VAL' AND DT_VALUE > 150" & _
             "GROUP BY DT_MEASURECOD"
    If rsExtractor.selectionfast(strSQL) Then
        If Not rsExtractor.IsEOF Then NumSupero10MinCO = rsExtractor.GetValue("Tot")
    End If
    
    objEngine.NumberValue_StringRange(0, "Y69") = NumTot10MinCO
    objEngine.NumberValue_StringRange(0, "Y70") = NumValidi10MinCO
    If NumValidi10MinCO Then
        objEngine.NumberValue_StringRange(0, "Y71") = NumSupero10MinCO
        objEngine.NumberValue_StringRange(0, "Y72") = NumSupero10MinCO / NumValidi10MinCO * 100
    Else
        objEngine.NumberValue_StringRange(0, "Y71") = 0
        objEngine.NumberValue_StringRange(0, "Y72") = 0
    End If
    
    Exit Sub
GestErrore:
    Call windasLog("Elabora10MinutiCO: " & Error(Err()))
End Sub

Private Sub IniziaReport()

    On Error GoTo GestErrore

    'Inizializzazione dell'oggetto per la creazione del report
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    Exit Sub
    
GestErrore:
    Call windasLog("IniziaReport: " & Err())

End Sub

Private Sub CalcolaPosizioniFoglio()

    On Error GoTo GestErrore
    
    COLONNA_IMPIANTO = (PARAMETRI_PER_PAGINA * COLONNE_PER_PARAMETRO) + COLONNA_PARAMETRI + 1
    RIGA_LIMITI_SINGOLI = RIGA_INTESTAZIONI + NR_RIGHE + 1
    RIGA_MINIMO = RIGA_LIMITI_SINGOLI + 2
    RIGA_MASSIMO = RIGA_MINIMO + 1
    RIGA_TOTALE_PERIODO = RIGA_MASSIMO + 2
    RIGA_LIMITI_PERIODO = RIGA_TOTALE_PERIODO + 1
    RIGA_CONTROLLO_97 = RIGA_MASSIMO + 2
    RIGA_LIMITI_97 = RIGA_CONTROLLO_97 + 1
    COLONNA_LEGENDA_NOTE = 1
    RIGA_LEGENDA_NOTE = RIGA_LIMITI_PERIODO + 5
    CELLA_ELAB_LEGGE = "M" + CStr(RIGA_LIMITI_PERIODO + 7)
    CELLA_O2_RIF = "M" + CStr(RIGA_LIMITI_PERIODO + 8)

    Exit Sub
    
GestErrore:
    Call windasLog("CalcolaPosizioniFoglio: " & Err())

End Sub

'Gestione della legenda note dinamica
Sub LegendaNote(ByVal report As String)
'report:
'G = Giornaliero
'GF = Giornaliero Flussi
'M = Mensile
'MF = Mensile Flussi
'T = Trimestrale anche Flussi
'A = Annuale anche Flussi

    On Error GoTo GestErrore

    Dim strLegenda As String
    Dim arrNote() As String
    Dim offset As Integer
    
    'Suppongo di scriverli tutti
    strLegenda = "(1);(2);(4);(5);(6);(8)"
    
    'Tolgo in base al tipo di report
    Select Case report
        Case "G"
            strLegenda = IIf(Loc("LINGUA") = "IT", strLegenda, Replace(strLegenda, "(1)", "  "))
        Case "GF", "MF"
            strLegenda = Replace(Replace(strLegenda, "(1);", ""), "(2);", "")
        Case "M"
            strLegenda = Replace(strLegenda, "(1);", "")
        Case "T", "A"
            strLegenda = Replace(Replace(strLegenda, "(1);", ""), "(6);", "")
        Case "AF"
            strLegenda = Replace(Replace(Replace(strLegenda, "(6);", ""), "(2);", ""), "(1);", "")
    End Select
    
    'Traduco per il Piemonte
    If IsPiemonte Then
        strLegenda = Replace(strLegenda, "(1)", "  ") 'Mi serve tenere la riga
        strLegenda = Replace(strLegenda, "(2)", "NS")
        strLegenda = Replace(strLegenda, "(4)", "NV")
        strLegenda = Replace(strLegenda, "(5)", "S")
        strLegenda = Replace(strLegenda, "(6)", "S")
    End If
    
    'Divido la stringa per ricavare quelle da scrivere
    arrNote = Split(strLegenda, ";")
    'offset = RIGA_LIMITI_PERIODO + 7
    offset = RIGA_LEGENDA_NOTE + 2
    'Scrivo le note nel foglio
    For z = offset To offset + UBound(arrNote)
        objEngine.StringValue_StringRange(0, "C" & CStr(z)) = arrNote(z - offset)
    Next z

    Exit Sub
    
GestErrore:
    Call windasLog("LegendaNote: " & Err())

End Sub

'Federica agosto 2017 - Estrazione minimo e massimo
Private Sub MinimoMassimo(ByVal tabella As String, ByVal dal As String, ByVal al As String, ByVal parametro As String)

    On Error GoTo GestErrore

    xMin = -9999
    xMax = -9999
    'Minimo e massimo
    strSQL = "SELECT MIN(DT_VALUE) AS Minimo, MAX(DT_VALUE) AS Massimo, COUNT(DT_VALUE) AS Righe " & _
             "FROM " & tabella & " WHERE DT_DATE between '" & dal & "' AND '" & al & "' " & _
             "AND DT_STATIONCODE  = '" & SelStation & "' and DT_MEASURECOD = '" & parametro & "' " & _
             "AND DT_VALIDFLAG IN('VAL', 'AUX', 'VAH')"
    With rsExtractor
        If (.selectionfast(strSQL)) Then
            If .GetValue("Righe") > 0 Then
                xMin = .GetValue("Minimo")
                xMax = .GetValue("Massimo")
            End If
        End If
    End With
    
    Exit Sub

GestErrore:
    Call windasLog("MinimoMassimo: " & Err())

End Sub

Private Sub MinimoMassimoFlussi(ByVal tabella As String, ByVal dal As String, ByVal al As String, ByVal parametro As String)

    On Error GoTo GestErrore

    xMin = -9999
    xMax = -9999
    'Minimo e massimo
    strSQL = "SELECT MIN(" & tabella & ".DT_FM) AS Minimo, MAX(" & tabella & ".DT_FM) AS Massimo, COUNT(" & tabella & ".DT_FM) AS Righe " & _
             "FROM " & tabella & " LEFT JOIN (SELECT DT_DATE, DT_VALUE, DT_VALIDFLAG " & _
                                             "FROM " & tabella & " WHERE DT_DATE BETWEEN '" & dal & "' AND '" & al & "' AND DT_STATIONCODE  = '" & SelStation & "' " & _
                                             "AND DT_MEASURECOD = '" & PortataPar & "' ) P ON " & tabella & ".DT_DATE = p.DT_DATE " & _
             "WHERE " & tabella & ".DT_DATE BETWEEN '" & dal & "' AND '" & al & "' AND " & tabella & ".DT_STATIONCODE  = '" & SelStation & "' " & _
             "AND " & tabella & ".DT_MEASURECOD = '" & parametro & "' " & _
             "AND " & tabella & ".DT_VALIDFLAG IN('VAL', 'AUX', 'VAH') and P.DT_VALUE <> -9999 AND P.DT_VALIDFLAG IN('VAL', 'AUX', 'VAH')"
    With rsExtractor
        If (.selectionfast(strSQL)) Then
            If .GetValue("Righe") > 0 Then
                xMin = .GetValue("Minimo")
                xMax = .GetValue("Massimo")
            End If
        End If
    End With
    
    Exit Sub
    
GestErrore:
    Call windasLog("MinimoMassimoFlussi: " & Err())

End Sub
Private Sub ScriviPartiFisse(ByVal sTipoReport As String, ByVal sTipoImpianto As String, ByVal sIntestazioneImpianto As String, _
ByVal sLimitiSingoli As String, ByVal sTotalePeriodo As String, ByVal sLimitiPeriodo As String, ByVal sData As String, ByVal sIntestaValori As String)
            
    On Error GoTo GestErrore
            
    'Nome impianot
    objEngine.StringValue_StringRange(0, CELLA_NOME_IMPIANTO) = NomeImpianto
    'Periodo
    objEngine.StringValue_StringRange(0, CELLA_DATA) = sData
    'Tipo Report
    objEngine.StringValue_StringRange(0, CELLA_TIPO_REPORT) = Loc(sTipoReport)
    'Stato/Ore impianto
    objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, COLONNA_IMPIANTO) = Loc(sTipoImpianto)
    'Valore per periodo
    objEngine.StringValue_CoordRange(0, RIGA_INTESTAZIONI, COLONNA_INTESTAZIONI) = Loc(sIntestaValori)
    'Colonne per parametro
    For z = COLONNA_PARAMETRI To (COLONNE_PER_PARAMETRO * PARAMETRI_PER_PAGINA) Step COLONNE_PER_PARAMETRO
        objEngine.StringValue_CoordRange(0, RIGA_INTESTAZIONI, z) = Loc("COL_NOTE")
        objEngine.StringValue_CoordRange(0, RIGA_INTESTAZIONI, z + 1) = Loc("COL_VAL")
        objEngine.StringValue_CoordRange(0, RIGA_INTESTAZIONI, z + 2) = Loc("COL_ID")
    Next z
    'Intestazione impianto
    objEngine.StringValue_CoordRange(0, RIGA_INTESTAZIONI, COLONNA_IMPIANTO) = Loc(sIntestazioneImpianto)
    'Limiti singoli
    objEngine.StringValue_CoordRange(0, RIGA_LIMITI_SINGOLI, COLONNA_INTESTAZIONI) = Loc(sLimitiSingoli)
    'Minimo
    objEngine.StringValue_CoordRange(0, RIGA_MINIMO, COLONNA_INTESTAZIONI) = Loc("ROW_MIN")
    'Massimo
    objEngine.StringValue_CoordRange(0, RIGA_MASSIMO, COLONNA_INTESTAZIONI) = Loc("ROW_MAX")
    'Media/somma periodo
    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_INTESTAZIONI) = Loc(sTotalePeriodo)
    'Limiti periodo
    objEngine.StringValue_CoordRange(0, RIGA_LIMITI_PERIODO, COLONNA_INTESTAZIONI) = Loc(sLimitiPeriodo)
    'Elaborazioni di Legge
    objEngine.StringValue_StringRange(0, CELLA_ELAB_LEGGE) = Loc("ELAB_DLGS")
    'Legenda annotazioni
    objEngine.StringValue_CoordRange(0, RIGA_LEGENDA_NOTE, COLONNA_LEGENDA_NOTE) = Loc("COL_NOTE") & ":"
    'Ossigeno di riferimento
    'Nicola 05/03/2015 Scrivo ossigeno rif
    If Val(O2_rif) > 0 Then
        objEngine.StringValue_StringRange(0, CELLA_O2_RIF) = Loc("STR_O2_1") & " " & O2_rif & Loc("STR_O2_2")
    ElseIf Val(O2_rif_s) > 0 Then
        objEngine.StringValue_StringRange(0, CELLA_O2_RIF) = Loc("STR_O2_1") & " " & O2_rif_s & Loc("STR_O2_2")
    End If
    '**************************************************

    Exit Sub
    
GestErrore:
    Call windasLog("ScriviPartiFisse:" & Err())

End Sub

Sub CreateReport()

    On Error GoTo GestErrore
     
    If ModelFileName <> "" Then
        vet = Split(ModelFileName, "\")
        
        Call CreateReportInizializza
        
        'Verifico la presenza del file con la tabella delle traduzioni per le lingue
        Call CaricaFileTraduzione
        
        Select Case vet(UBound(vet))  'ultima posizione array
            '*** ANNUALI ***
            Case "Wtf_ModelloAnnuale.xls":
                Call Annuale(NumMedieOrario, MinimoOreMese, False)
            Case "Wtf_ModelloAnnuale_FM.xls":
                Call AnnualeFM(False)
            Case "Wtf_ModelloAnnualeSemiorario.xls":
                Call Annuale(NumMedieSemiorario, MinimoSemioreMese, True)
            Case "Wtf_ModelloAnnualeSemiorario_FM.xls":
                Call AnnualeFM(True)
                
            '*** TRIMESTRALI ***
            Case "Wtf_ModelloTrimestrale.xls":
                Call Trimestrale(MinimoOreMese, False)
            Case "Wtf_ModelloTrimestrale_FM.xls":
                Call TrimestraleFM(False)
            
            '*** MENSILI ***
            Case "Wtf_ModelloMensile.xls":
                Call Mensile(NumMedieOrario, MinimoOreGiorno, MinimoOreMese, False)
            Case "Wtf_ModelloMensile_FM.xls":
                Call MensileFM
            Case "Wtf_ModelloMensileSemiorario.xls":
                Call Mensile(NumMedieSemiorario, MinimoSemioreGiorno, MinimoSemioreMese, True)
            Case "Wtf_ModelloMensileSemiorario_FM.xls":
                Call MensileFM(True)
                        
            '*** ORARI/SEMIORARI ***
            Case "Wtf_ModelloOrario.xls":
                Call Giornaliero(True, MinimoOreGiorno)
            Case "Wtf_ModelloOrario_FM.xls":
                Call GiornalieroFM(False)
            Case "Wtf_ModelloOrario_FM2.xls":   'normal transitorio + transitorio
                Call GiornalieroFM(True)
            Case "Wtf_ModelloSemiorario.xls":
                Call Giornaliero(True, MinimoSemioreGiorno, True)
            Case "Wtf_ModelloSemiorario_FM.xls":
                Call GiornalieroFM(False, True)
            Case "Wtf_ModelloOrario_TQ.xls":    'medie orarie grezze
                Call Giornaliero(False, MinimoOreGiorno)
            
            '*** QAL2 ***
            Case "Wtf_ModelloQAL2.xls":
                Call QAL2
            Case "Wtf_ModelloQAL2Semiorario.xls":
                Call QAL2
            
            '*** QAL3 ***
            Case "Wtf_ModelloQAL3Cusum.xls":
                Call QAL3
            
            '*** ALTRI ***
            Case "Wtf_ModelloDL133.xls":
                Call StatisticheDlgs133
                        
            '*** EXPORT ***
            Case "Wtf_ModelloExportSAD.xls"
                Call ExportSAD
            Case "Wtf_ModelloExportOrario.xls"
                Call ExportOrario
            Case "Wtf_ModelloExportOrarioGruppi.xls":
                Call ExportOrarioGruppi
        
            'luca novembre 2018
            Case "Wtf_ModelloConteggio.xls":
                Call Totalizzazioni("093")
        End Select
    End If
    
    Exit Sub
    
GestErrore:
    Call windasLog("CreatReport: " & Error(Err))
    'Stop
    Resume Next

End Sub
'luca 20/10/2015
Public Sub ExportSAD()
  
    'Dim CheckDate As Date
    Dim StartRiga As Integer
    Dim iSheet As Integer
    
    iSheet = 0

    '************************************************************
    '*                                                          *
    '************************************************************
    On Local Error GoTo GestErrore
    With rsExtractor
    
        
        '**************************************************************************************
        '*                                lettura file SAD                                    *
        '**************************************************************************************
        Dim Nome_File As String
        Dim NomeFile As String
        Dim PathARPA As String
        Dim ll As Integer
        Dim colonne() As String
        Dim Dati() As String
        Dim sLinea As String
        Dim ndx As Integer
        
        '***** Nome file *****
        .selectionfast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='4343_FILE'"
        Nome_File = .GetValue("gt_value")

        '***** Path file arpa *****
        .selectionfast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='4343_PATH'"
        PathARPA = .GetValue("gt_value")

        '**** punta alla prima riga dati del foglio EXCEL
        StartRiga = 1

        CheckDate = StartDate
        
        '***** Nome file SAD *****
        NomeFile = Nome_File & "_" & Format(CheckDate, "yyyymmdd") & ".SAD"

        ll = FreeFile

        '***** Se il file esiste *****
        If Dir(PathARPA & "\" & NomeFile, vbArchive + vbHidden) <> "" Then

            '***** Apertura del file *****
            Open PathARPA & "\" & NomeFile For Input As #ll
            
            Do While Not EOF(ll)
              Line Input #ll, sLinea
              Dati = Split(sLinea, Chr(9))
              For ndx = 0 To UBound(Dati)
                '***** Sostituzione flag di validità per linee guida Regione Piemonte *****
                Call SostituisciFlagPiemonte(Dati(ndx))
                objEngine.StringValue_StringRange(iSheet, CoordToString(StartRiga, ndx + 1)) = Dati(ndx)
              Next ndx
              StartRiga = StartRiga + 1
            Loop
            Close (ll)

        End If
        
    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Debug.Print Error(Err)
    Call windasLog("ExportSAD: " & Error(Err))
    Resume Next

End Sub

'luca maggio 2017
Public Sub ExportOrario()
  
    Dim MaxPar As Integer
    'Dim CurRow As Integer
    Dim DataCorrente As Date
    Dim Validflag As String
    
    Const RigaIniziale = 7

    On Error GoTo GestErrore
    
    CurRow = RigaIniziale
    
    xStartRange = 8
    yStartRange = 2
    xEndRange = 8
    yEndRange = 32
    yDestTopLeft = 2
        
    With rsExtractor
          
        objEngine.StringValue_StringRange(0, "B2") = NomeImpianto
        
        objEngine.StringValue_StringRange(0, "AB2") = Loc("DA_DATA") & " " & Format(StartDate, "dd/mm/yyyy") & "  " & Loc("A_DATA") & " " & Format(EndDate, "dd/mm/yyyy")
            
        ListaParametri = Split(SelParam, ",")
        MaxPar = UBound(ListaParametri)
                    
        strSQL = "SELECT * FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_DATE >= " & String_StartDate_DB & " AND DT_DATE <= " & String_EndDate_DB & " ORDER BY DT_DATE,DT_HOUR"
        .selectionfast (strSQL)
                            
        .m_Filter = "DT_MEASURECOD = '" & NFunzPar & "'"
            
        Do While Not .IsEOF
            'Formatto la riga successiva come quella attuale
            Call objEngine.RangeCopyTo(CoordToString(xStartRange, yStartRange), CoordToString(xEndRange, yEndRange), CoordToString(xEndRange + 1, yDestTopLeft))
            'aggiorno riga da cui copio la formattazione
            xStartRange = xStartRange + 1
            xEndRange = xEndRange + 1
            
            'scrivo data/ora
            DataCorrente = CreateDateFromDB(Trim(.GetValue("dt_date")))
            objEngine.StringValue_CoordRange(0, CurRow, 1) = CStr(DataCorrente) + " - " + Trim(.GetValue("DT_HOUR"))
            
            'scrivo stato impianto
            If .GetValue("DT_CUSTOM1") <> "-9999" Then
                objEngine.StringValue_CoordRange(0, CurRow, 2) = GetStsImpDescription(CInt(.GetValue("DT_CUSTOM1")))
            End If
            
            'scrivo codice impianto
            objEngine.StringValue_CoordRange(0, CurRow, 3) = RicavaCodiceImpianto(CInt(Trim(Left(SelStation, 1))), False)
            
            CurRow = CurRow + 1
            .movenext
        Loop
            
        '***** Ciclo per parametro ****************
        For i = 0 To MaxPar
            'primi 2 parametri (CO e NOx) gestisco valore, validità ed ID x3 (TQ, NORM, FLUSSO)
            If i < 2 Then
                objEngine.StringValue_CoordRange(0, 5, (9 * i) + 4) = CodPar(i, 1) & " TQ (" & CodPar(i, 2) & ")"
                objEngine.StringValue_CoordRange(0, 5, (9 * i) + 7) = CodPar(i, 1) & " NORM (" & CodPar(i, 2) & ")"
                objEngine.StringValue_CoordRange(0, 5, (9 * i) + 10) = CodPar(i, 1) & " FLUSSO (Kg/h)"
            'terzo parametro (O2) gestisco valore, validità e ID
            ElseIf i = 2 Then
                objEngine.StringValue_CoordRange(0, 5, 22) = CodPar(i, 1) & " (" & CodPar(i, 2) & ")"
            'altri parametri gestisco valore
            Else
                objEngine.StringValue_CoordRange(0, 5, i + 22) = CodPar(i, 1) & " (" & CodPar(i, 2) & ")"
            End If
            
            '***** Filtro per parametro *****
            .m_Filter = "DT_MEASURECOD = '" & CodPar(i, 0) & "'"
            If Not .IsEOF Then
                .MoveFirst
            End If
            
            CurRow = RigaIniziale
            
            Do While Not .IsEOF
                'primi 2 parametri (CO e NOx) gestisco valore, validità ed ID x3 (TQ, NORM, FLUSSO)
                If i < 2 Then
                    'TAL QUALE
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 4) = Format(.GetValue("DT_VALUETQ"), "0.0")
                    Validflag = Trim(.GetValue("DT_VALIDFLAG_TQ"))
                    Call SostituisciFlagPiemonte(Validflag)
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 5) = Validflag
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 6) = Format(Val(.GetValue("DT_NR_TQ")) / 720 * 100, "0.0")
                    'NORM
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 7) = Format(.GetValue("DT_VALUE"), "0.0")
                    Validflag = Trim(.GetValue("DT_VALIDFLAG"))
                    Call SostituisciFlagPiemonte(Validflag)
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 8) = Validflag
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 9) = Format(Val(.GetValue("DT_NR")) / 720 * 100, "0.0")
                    'FLUSSO
                    Dim statusPortata As String
                    Dim NrPortata As Integer
                    strSQL = "SELECT * FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_DATE = '" & Trim(.GetValue("DT_DATE")) & "' AND DT_HOUR = '" & Trim(.GetValue("DT_HOUR")) & "' AND DT_MEASURECOD = '" & PortataPar & "'"
                    rsNfunz.selectionfast strSQL
                    statusPortata = CStr(Trim(rsNfunz.GetValue("DT_VALIDFLAG")))
                    NrPortata = Val(Trim(rsNfunz.GetValue("DT_NR")))
                    
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 10) = Format(.GetValue("DT_FM"), "0.0")
                    'se valore di concentrazione e portata sono validi il flusso è valido, altrimenti no
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 11) = IIf((.GetValue("DT_VALIDFLAG") = "VAL" Or .GetValue("DT_VALIDFLAG") = "AUX") And (statusPortata = "VAL" Or statusPortata = "AUX"), "V", "AL")
                    'come ID prendo quello più basso tra l'ID della misura di concentrazione e quella di portata
                    objEngine.StringValue_CoordRange(0, CurRow, (9 * i) + 12) = Format(IIf(Val(Trim(.GetValue("DT_NR"))) < NrPortata, Val(Trim(.GetValue("DT_NR"))) / 720 * 100, NrPortata / 720 * 100), "0.0")
                'terzo parametro (O2) gestisco valore, validità e ID
                ElseIf i = 2 Then
                    objEngine.StringValue_CoordRange(0, CurRow, 22) = Format(.GetValue("DT_VALUE"), "0.0")
                    Validflag = Trim(.GetValue("DT_VALIDFLAG"))
                    Call SostituisciFlagPiemonte(Validflag)
                    objEngine.StringValue_CoordRange(0, CurRow, 23) = Validflag
                    objEngine.StringValue_CoordRange(0, CurRow, 24) = Format(Val(Replace(.GetValue("DT_NR"), ",", ".")) / 720 * 100, "0.0")
                'altri parametri gestisco valore
                Else
                    objEngine.StringValue_CoordRange(0, CurRow, i + 22) = Format(.GetValue("DT_VALUE"), "0.0")
                End If
                CurRow = CurRow + 1
                .movenext
            Loop
        Next i
        .m_Filter = ""
    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    Set rsNfunz = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("ExportOrario: " & Error(Err))

End Sub
'luca maggio 2017
Public Sub ExportOrarioGruppi()
  
    Dim MaxPar As Integer
    'Dim CurRow As Integer
    Dim DataCorrente As Date
    Dim Validflag As String
    
    Const RigaIniziale = 7

    On Error GoTo GestErrore
    
    CurRow = RigaIniziale
    
    xStartRange = 8
    yStartRange = 2
    xEndRange = 8
    yEndRange = 9
    yDestTopLeft = 2
        
    With rsExtractor
          
        objEngine.StringValue_StringRange(0, "B2") = "Alba Power - Gruppo " & IIf(CInt(Left(Trim(SelStation), 1) = 1), "1", "2")
        
        objEngine.StringValue_StringRange(0, "G2") = Loc("DA_DATA") & " " & Format(StartDate, "dd/mm/yyyy") & "  " & Loc("A_DATA") & " " & Format(EndDate, "dd/mm/yyyy")
            
        ListaParametri = Split(SelParam, ",")
        MaxPar = UBound(ListaParametri)
                    
        strSQL = "SELECT * FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_DATE >= " & String_StartDate_DB & " AND DT_DATE <= " & String_EndDate_DB & " ORDER BY DT_DATE,DT_HOUR"
        .selectionfast (strSQL)
        
        'filtro per stato impianto per ottenere il riferimento reale delle ore effettive (anche se non serve l'informazione di stato impianto)
        .m_Filter = "DT_MEASURECOD = '" & NFunzPar & "'"
            
        Do While Not .IsEOF
            'Formatto la riga successiva come quella attuale
            Call objEngine.RangeCopyTo(CoordToString(xStartRange, yStartRange), CoordToString(xEndRange, yEndRange), CoordToString(xEndRange + 1, yDestTopLeft))
            'aggiorno riga da cui copio la formattazione
            xStartRange = xStartRange + 1
            xEndRange = xEndRange + 1
            
            'scrivo data/ora
            DataCorrente = CreateDateFromDB(Trim(.GetValue("dt_date")))
            objEngine.StringValue_CoordRange(0, CurRow, 1) = CStr(DataCorrente) + " - " + Trim(.GetValue("DT_HOUR"))
            
            'scrivo codice impianto
            objEngine.StringValue_CoordRange(0, CurRow, 2) = RicavaCodiceImpianto(CInt(Trim(Left(SelStation, 1))), True)
            
            CurRow = CurRow + 1
            .movenext
        Loop
            
        '***** Ciclo per parametro ****************
        For i = 0 To MaxPar
        
            objEngine.StringValue_CoordRange(0, 5, i + 3) = CodPar(i, 1) & " (" & CodPar(i, 2) & ")"
            
            '***** Filtro per parametro *****
            .m_Filter = "DT_MEASURECOD = '" & CodPar(i, 0) & "'"
            If Not .IsEOF Then
                .MoveFirst
            End If
            
            CurRow = RigaIniziale
            
            Do While Not .IsEOF
                objEngine.StringValue_CoordRange(0, CurRow, i + 3) = Format(.GetValue("DT_VALUE"), "0.0")
                CurRow = CurRow + 1
                .movenext
            Loop
        Next i
        .m_Filter = ""
    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("ExportOrarioGruppi: " & Error(Err))

End Sub
Function CalcolaRigaSemiora(ByVal semiora As Integer) As Integer

    On Error GoTo GestErrore

    If semiora <> 0 Then
        CalcolaRigaSemiora = 1
    Else
        CalcolaRigaSemiora = 0
    End If
    
    Exit Function
    
GestErrore:
    Call windasLog("CalcolaRigaSemiora: " & Err())
    
End Function

'Nicolò aprile 2013 Check per vedere se il codice interno report si riferisce ai flussi di massa
Public Function IsFlussi(tipo As String) As Boolean
    IsFlussi = (Len(tipo) > 2) And (Mid(tipo, 2, 1) = "f")
End Function
Public Sub Giornaliero(isNorm As Boolean, ByVal MinimoOre As Integer, Optional ByVal isSemiore As Boolean = False)
    
    Dim ContaValidi(50, 1)
    Dim SommaValidi(50, 1)
    Dim Value As String
    Dim Validflag As String
    Dim StatoImpianto As String
    Dim nn As Integer
    Dim campoNumCampioni As String 'Nicolò Settembre 2105
    'luca 25/01/2016
    Dim ValidFlagPiemonte As String
    Dim notaDaScrivere As String 'Federica agosto 2017 - per la gestione di lingua e Piemonte
    'Federica settembre 2017 - per la gestione delle semiore
    Dim LimiteDaConfrontare As Double
    Dim NrCampioniTotali As Integer
    Dim indiceOra As Integer
    
    On Error GoTo GestErrore
    
    'Edoardo 26/10/2016 diferenziazione variabili in base al tipo di report
    'normalizzato o grezzo
    If isNorm Then
        Value = "dt_value"
        Validflag = "dt_validflag"
        campoNumCampioni = "DT_NR" 'Nicolò Settembre 2015
    Else
        Value = "dt_valuetq"
        Validflag = "dt_validflag_tq"
        campoNumCampioni = "DT_NR_TQ" 'Nicolò Settembre 2015
    End If
    StatoImpianto = "DT_CUSTOM1"
        
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = IIf(isSemiore, NumMedieSemiorario, NumMedieOrario) + 1
    NrCampioniTotali = IIf(isSemiore, 360, 720)
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    '***** Scrittura delle parti fisse del foglio *****
    If isSemiore Then
        RIGA_LIMITI_PERIODO = RIGA_LIMITI_PERIODO + 3
        RIGA_TOTALE_PERIODO = RIGA_TOTALE_PERIODO + 3
    
        Call ScriviPartiFisse("RPT_GIORNO_S", "COL_IMPIANTO", "TIPO_STATOIMP", "ROW_LIMITI_SEMIORA", "ROW_MEDIA_GIORNO", _
            "ROW_LIMITI_GIORNO", Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy"), "COL_ORA")
    Else
        Call ScriviPartiFisse("RPT_GIORNO", "COL_IMPIANTO", "TIPO_STATOIMP", "ROW_LIMITI_ORA", "ROW_MEDIA_GIORNO", _
            "ROW_LIMITI_GIORNO", Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy"), "COL_ORA")
    End If
    Call LegendaNote("G")
    For z = RIGA_LEGENDA_NOTE + 2 To RIGA_LEGENDA_NOTE + 7
        If (z = RIGA_LIMITI_PERIODO + 7) Then 'Gestione note solo Italia
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = IIf(Loc("LINGUA") = "IT", Loc(IIf(isSemiore, "NOTE_OS_", "NOTE_O_") & CStr(z)), "")
        Else
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc(IIf(isSemiore, "NOTE_OS_", "NOTE_O_") & CStr(z))
        End If
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** Azzeramento matrici *****
    'For z = 0 To 23
    For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
        NFunzValue(z) = -9999
    Next
        
    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        'Federica luglio 2017 - gestione della traduzione del nome parametro in LocTable.ini
        'objEngine.StringValue_CoordRange(0, 5, (3 * i) + 2) = CodPar(i, 1) + " (" + CodPar(i, 2) + ")"
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" + CodPar(i, 2) + ")"
        If isSemiore Then
            If LimiteGiorno(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_PERIODO, CurCol + POS_COL_VALORE) = LimiteGiorno(i)
            If LimiteSemiora100(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteSemiora100(i)
            If LimiteSemiora97(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_97, CurCol + POS_COL_VALORE) = LimiteSemiora97(i)
        Else
            If LimiteGiorno(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_PERIODO, CurCol + POS_COL_VALORE) = LimiteGiorno(i)
            If LimiteOra(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteOra(i)
        End If
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    With rsExtractor
        'luca dicembre 2014 aggiorno campi di lettura medie orarie
        '***** Estrazione medie orarie *****
        strSQL = "SELECT * FROM " & SelTable & " WHERE " & _
                " DT_STATIONCODE = '" & SelStation & "' AND DT_DATE = '" & String_StartDate_DB & "' ORDER BY DT_HOUR"
          
        If (.selectionfast(strSQL)) Then
            '***** Controllo se impianto è in normale funzionamento *****
            '***** Filtro per stato impianto *****
            .m_Filter = "DT_MEASURECOD = '" & NFunzPar & "'"
            
            '***** Caricamento in matrice dei valori di stato impianto *****
            Do While Not rsExtractor.IsEOF
                'luca dicembre 2014 aggiorno campo lettura stato impianto
                indiceOra = Left(.GetValue("DT_HOUR"), 2)
                If isSemiore Then indiceOra = indiceOra * 2 + CalcolaRigaSemiora(Right(.GetValue("DT_HOUR"), 2))
                NFunzValue(indiceOra) = .GetValue(StatoImpianto)
                
                .movenext
            Loop
            
            'luca gennaio 2015 modifico gestendo tutti i codici stato impianto possibili (ABS UDINE)
            '***** Inserimento stato impianto orario *****
            '***** Ciclo per ore *****
            'For z = 0 To 23
            For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
                'Nicolò Agosto 2015 rendo più flessibile la gestione dei testi per lo stato impianto
                If CDbl(NFunzValue(z)) <> -9999 Then
                    If CDbl(NFunzValue(z)) = 30 Then
                        OreNormFunz = OreNormFunz + 1
                    End If
                    CurRow = RIGA_INTESTAZIONI + 1 + z
                    objEngine.StringValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = GetStsImpDescription(CInt(NFunzValue(z)))
                End If
            Next
            
            '***** Ciclo per parametro *****
            For i = 0 To MaxPar
                '***** Filtro per parametro *****
                .m_Filter = "DT_MEASURECOD = '" & CodPar(i, nn) & "'"
                xMin = 999999
                xMax = -999999
                
                Do While Not rsExtractor.IsEOF
                    '***** Coordinate dove scrivere *****
                    indiceOra = Left(.GetValue("DT_HOUR"), 2)
                    If isSemiore Then indiceOra = indiceOra * 2 + CalcolaRigaSemiora(Right(.GetValue("DT_HOUR"), 2))
                    CurRow = indiceOra + RIGA_INTESTAZIONI + 1
                    CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
                    
                    'luca 25/01/2016 validflag
                    ValidFlagPiemonte = .GetValue(Validflag)
                    Call SostituisciFlagPiemonte(ValidFlagPiemonte)
                    objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = ValidFlagPiemonte
                                        
                    'luca dicembre 2014 aggiorno campo di lettura valore
                    If Trim(.GetValue(Value)) = "-9999" Then
                        '***** Assenza registrazione medie *****
                        'Federica agosto 2017 - Gestisco se lingua <> da IT e Piemonte
                        notaDaScrivere = IIf(Loc("LINGUA") = "IT", IIf(IsPiemonte, "", "'(1)"), "'(2)")
                        objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = notaDaScrivere
                    Else
                        '***** Media oraria *****
                        'luca dicembre 2014 aggiorno campo di lettura valore
                        objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = .GetValue(Value)
                        
                        'luca 25/01/2016 disabilito gestione media invalida in quanto già gestita da validflag piemonte
                        If Not IsPiemonte And InStr(strValid, Trim(.GetValue(Validflag))) = 0 Then
                            '***** Media invalida *****
                            objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = "'(4)"
                        ElseIf InStr(strValid, Trim(.GetValue(Validflag))) > 0 Then
                            '***** Se lo stato impianto è "In Marcia" *****
                            If CDbl(NFunzValue(indiceOra)) <> -9999 Then
                                If CDbl(NFunzValue(indiceOra)) = 30 Then
                                
                                    '***** Aggiornamento della sommatoria e del contatore per il calcolo della media giornaliera *****
                                    ContaValidi(i, nn) = ContaValidi(i, nn) + 1
                                    'luca dicembre 2014 aggiorno campo di lettura valore
                                    SommaValidi(i, nn) = SommaValidi(i, nn) + .GetValue(Value)
                                    
                                    If isSemiore Then
                                        LimiteDaConfrontare = LimiteSemiora100(i)
                                    Else
                                        LimiteDaConfrontare = LimiteOra(i)
                                    End If
                                    
                                    '***** Controllo rispetto limite orario *****
                                    If LimiteDaConfrontare > 0 Then
                                        'luca dicembre 2014 aggiorno campo di lettura valore
                                        If CDbl(.GetValue(Value)) > CDbl(LimiteDaConfrontare) Then
                                            'luca 25/01/2016
                                            objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "S", "'(5)")
                                        End If
                                    End If
                                    
                                    '***** Statistica Media oraria minima *****
                                    If CDbl(.GetValue(Value)) < CDbl(xMin) Then xMin = .GetValue(Value)
                                    '***** Statistica Media oraria massima *****
                                    If CDbl(.GetValue(Value)) > CDbl(xMax) Then xMax = .GetValue(Value)
                                End If
                            End If
                        End If
                    End If
                
                    '***** Inserimento ID% orario *****
                    objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = .GetValue(campoNumCampioni) / NrCampioniTotali * 100 'Nicolò Settembre 2015 passo a dt_nr

                    .movenext
                Loop
                
                '***** Inserimento valore orario minimo *****
                If xMin <> 999999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i) + POS_COL_VALORE) = xMin
                
                '***** Inserimento valore orario massimo *****
                If xMax <> -999999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i) + POS_COL_VALORE) = xMax
            Next
            .m_Filter = ""
        End If
            
        '***** Ciclo per parametro *****
        For i = 0 To MaxPar
            CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
            If Len(CodPar(i, nn)) > 0 Then
                '***** Se ho il minimo di normal funzionamento *****
                If CDbl(OreNormFunz) >= MinimoOre Then
                    '***** Se c'è almeno una media oraria valida *****
                    If ContaValidi(i, nn) > 0 Then
                        '***** Media giornaliera *****
                        objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = SommaValidi(i, nn) / ContaValidi(i, nn)
                        'luca 25/01/2016 media giornaliera valida - piemonte
                        If IsPiemonte Then objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + 2) = "V"
                        
                        '***** Status di validità media giornaliera *****
                        '***** Media giornaliera valida solo con ID% >= 70% *****
                        If CDbl(ContaValidi(i, nn) / OreNormFunz * 100) < 70 Then
                            'luca 25/01/2016
                            objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "NV", "'(4)")
                        Else
                            '***** Controllo rispetto limite giornaliero *****
                            If LimiteGiorno(i) > 0 Then
                                If CDbl(SommaValidi(i, nn) / ContaValidi(i, nn)) > CDbl(LimiteGiorno(i)) Then
                                    'luca 25/01/2016
                                    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "S", "'(6)")
                                End If
                            End If
                        End If
                        
                        '***** ID% media giornaliera *****
                        objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = ContaValidi(i, nn) / OreNormFunz * 100
                    Else
                        '***** Status di validità media giornaliera - non calcolata *****
                        'luca 25/01/2016
                        objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "NV", "'(4)")
                        '***** ID% media giornaliera *****
                        objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = 0
                    End If
                    
                Else
                    '***** Status di validità media giornaliera - non significativa*****
                    'luca 25/01/2016
                    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "NS", "'(2)")
                    '***** ID% media giornaliera *****
                    objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = 0
                End If
            End If
        Next
            
        '****** Inserimento numero ore di normale funzionamento giornaliere ******
        objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = OreNormFunz
    
        If isSemiore Then
            
            For i = 0 To MaxPar
                CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
                '****** Controllo 97% medie semiore nell'anno ******
                objEngine.NumberValue_CoordRange(0, RIGA_CONTROLLO_97, CurCol + POS_COL_VALORE) = Controllo97(i, year(StartDate))
                '****** Controllo 10 minuti CO ******
                If Abilita10minCO Then
                    'Call Elabora10MinutiCO("CO")
                    Call Elabora10MinutiCO("CO_s") 'Personalizzato Polar Piscinas
                End If
            Next i
        End If
    
        '***** Inserimento allarmi con riconoscimento *****
        objEngine.StringValue_StringRange(1, "C2") = NomeImpianto
        objEngine.StringValue_StringRange(1, "G2") = Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy")
        objEngine.StringValue_StringRange(1, "C4") = Loc("RPT_GIORNO_ALLARMI")
        objEngine.StringValue_StringRange(1, "C6") = Loc("COL_ALLARMI")
        objEngine.StringValue_StringRange(1, "D5") = Loc("COL_INIZIO")
        objEngine.StringValue_StringRange(1, "F5") = Loc("COL_FINE")
        objEngine.StringValue_StringRange(1, "D6") = Loc("COL_DATA")
        objEngine.StringValue_StringRange(1, "F6") = Loc("COL_DATA")
        objEngine.StringValue_StringRange(1, "E6") = Loc("COL_ORA")
        objEngine.StringValue_StringRange(1, "G6") = Loc("COL_ORA")
        strSQL = "SELECT * FROM WDS_ALARM WHERE AL_STATION = '" & SelStation & "' AND AL_DATE = '" & String_StartDate_DB & "' ORDER BY AL_DATE, AL_HOUR, AL_DESCRIPTION" 'Nicolò ordino anche per data e ora
                     
        If (.selectionfast(strSQL)) Then
            Do While Not rsExtractor.IsEOF
                 '***** Allarme *****
                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 2) = .GetValue("AL_DESCRIPTION") & " " & .GetValue("AL_STATUSDESC")
                 
                 '*************************** inizio allarme **************************
                 '***** Data *****
                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 3) = Right(.GetValue("AL_DATE"), 2) & "/" & Mid(.GetValue("AL_DATE"), 5, 2) & "/" & Left(.GetValue("AL_DATE"), 4)
                 '***** Ora *****
                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 4) = .GetValue("AL_HOUR")
     
                 '*************************** rientro allarme **************************
                 '***** Data *****
                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 5) = Right(.GetValue("AL_DATE2"), 2) & "/" & Mid(.GetValue("AL_DATE2"), 5, 2) & "/" & Left(.GetValue("AL_DATE2"), 4)
                 '***** Ora *****
                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 6) = .GetValue("AL_HOUR2")
    
                 RigaAlarm = RigaAlarm + 1
                 .movenext
            Loop
        End If

    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("Giornaliero: " & Error(Err))
    Resume Next
    
End Sub

Sub CreateReportInizializza()
    
    On Error GoTo GestErrore
    
    '******* NOTA BENE: Parametri da impostare per ogni impianto ******
      
    'Fattore correttivo portata, utile quando abbiamo ad esempio la portata in KNm3.
    FattorePortata = 1
    
    'Nicola da implementare
    maxNumeroAcquisizioni = 720
    
    'luca 04/06/2015 inizializzo RigaAlarm (riga iniziale per scrittura allarmi report giornaliero)
    'la setto uguale a zero poi nel corpo del report aggiunge il numero di righe corretto
    RigaAlarm = 0
    
    'luca 04/06/2015 setto le variabili per la gestione orario/semiorario
    NumMedieOrario = 23
    NumMedieSemiorario = 47
    MinimoOreGiorno = 6
    MinimoSemioreGiorno = 12
    MinimoOreMese = 144
    MinimoSemioreMese = 288
    MinimoOreTrimestre = MinimoOreMese * 3 'Federica luglio 2017 - per gestione report trimestrali
    
    'Federica luglio 2017 - posizionamenti standard nei modelli
    offset_riga = 7
    offset_colonna = 2
    '**********************************************************
    
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    Set rsExtractor = CreateObject("AttimoFwk.CData")
    Set rsNfunz = CreateObject("AttimoFwk.CData")
    Set rsValMin = CreateObject("AttimoFwk.CData")
    
    rsExtractor.Err_Activate = True
    rsExtractor.SetMessages False
    rsNfunz.Err_Activate = True
    rsNfunz.SetMessages False

    With rsExtractor
        Select Case DbType
            Case "MYSQL5"
                .SetDBType .Conn_MYSQL5
                rsNfunz.SetDBType .Conn_MYSQL5
            Case "MYSQL"
                .SetDBType .Conn_MYSQL
                rsNfunz.SetDBType .Conn_MYSQL
            Case "ORACLE"
                .SetDBType .Conn_Oracle
                rsNfunz.SetDBType .Conn_Oracle
            Case "SQL"
                .SetDBType .Conn_SQL
                rsNfunz.SetDBType .Conn_SQL
        End Select
        
        '***** Database, utente e password *****
        .SetDatabase DbDatabase, DbUser, DbPassword
        rsNfunz.SetDatabase DbDatabase, DbUser, DbPassword
        '***** Server ******
        .SetServer DbServer
        rsNfunz.SetServer DbServer
        '***** DB version ******
        .SetDbVersion DbVersion
        rsNfunz.SetDbVersion DbVersion
        
        '***** Stazione *****
        SelStation = Stations
        '***** parametri *****
        SelParam = Param
        '***** Tabella dati *****
        SelTable = Table
        
        
        '***** Data inizio e fine *****
        String_StartDate_DB = CreateDateForDB(StartDate)
        String_EndDate_DB = CreateDateForDB(EndDate)
        
        ListaParametri = Split(SelParam, ",")
        MaxPar = UBound(ListaParametri)
        'Ciclo per parametro
        For ndx = 0 To MaxPar
            
            CodPar(ndx, 0) = ListaParametri(ndx)
            
            '***** Estrazione della descrizione par. *****
            strSQL = "SELECT GT_Description FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & CodPar(ndx, 0) & "' AND GT_Description IS NOT NULL"
            .selectionfast (CStr(strSQL))
            If Not .IsEOF Then CodPar(ndx, 1) = .GetValue("GT_Description")

            '***** Estrazione dell'unità di misura *****
            strSQL = "SELECT GT_Str1 FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & CodPar(ndx, 0) & "' AND GT_Str1 IS NOT NULL"
            .selectionfast (CStr(strSQL))
            If Not .IsEOF Then CodPar(ndx, 2) = .GetValue("GT_Str1")
            
            '***** Estrazione della seconda unità di misura *****
            strSQL = "SELECT GT_Str2 FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & CodPar(ndx, 0) & "' AND GT_Str1 IS NOT NULL"
            .selectionfast (CStr(strSQL))
            If Not .IsEOF Then CodPar(ndx, 3) = .GetValue("GT_Str2")
            
            '********** estrae il numero di decimali da riportare ***********
            strSQL = "SELECT GT_NumInt FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = " & .ParSQLStr(CodPar(ndx, 0)) & " AND GT_NumInt IS NOT NULL"
            If (.selectionfast(strSQL)) Then
                decimali = Val(.GetValue("GT_NumInt"))
            Else
                decimali = 0
            End If
            CodPar(ndx, 4) = Format(decimali)

        Next ndx

        'Nicola Estraggo e scrivo la descrizione dello stato impianto da DB
        '************ estrae il nome completo della stazione ************
        strSQL = "SELECT GT_Description FROM wds_GenTab WHERE GT_Type = 'SYSTEMS'"
        .selectionfast (strSQL)
        If (Not .IsEOF) Then
            NomeImpianto = .GetValue("GT_Description")
        Else
            NomeImpianto = ""
        End If
        NomeImpianto = NomeImpianto & " "
        strSQL = "SELECT GT_Description FROM wds_GenTab WHERE GT_Type = 'STATIONS' AND GT_Code = " & .ParSQLStr(SelStation)
        .selectionfast (strSQL)
        If (Not .IsEOF) Then
            NomeImpianto = Trim(NomeImpianto & .GetValue("GT_Description"))
        Else
            NomeImpianto = Trim(NomeImpianto & SelStation)
        End If
        
         'STATO IMPIANTO
        strSQL = "SELECT c2 FROM was_measures WHERE LEFT(c23, 3) =(SELECT CC_Text FROM was_config WHERE CC_Code=50 and CC_Stationcode='" & SelStation & "') and CM_StationCode='" & SelStation & "'"
        .selectionfast (strSQL)
        If Not .IsEOF And Trim(.GetValue("c2")) <> "" Then
            NFunzPar = Trim(.GetValue("c2"))
        Else
            NFunzPar = "IMPIANTO"
        End If
        
        'PORTATA
        strSQL = "SELECT c2 FROM was_measures WHERE LEFT(c23, 3) =(SELECT CC_Text FROM was_config WHERE CC_Code=51 and CC_Stationcode='" & SelStation & "') and CM_StationCode='" & SelStation & "'"
        .selectionfast (strSQL)
        If Not .IsEOF And Trim(.GetValue("c2")) <> "" Then
            PortataPar = Trim(.GetValue("c2"))
        Else
            PortataPar = "QFUMI"
        End If
          
        'O2 RIF
        strSQL = "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & SelStation & "' and CC_CODE='0'"
        .selectionfast (strSQL)
        If Not .IsEOF Then
            O2_rif = .GetValue("cc_value")
        End If
        '***** Estrazione e Scrittura dell'ossigeno di riferimento con assetto medie semiorarie*****
        'Federica settembre 2017: Personalizzato Polar Piscinas
        '.selectionFast "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & SelStation & "' and CC_CODE='1202'"
        .selectionfast "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & SelStation & "' and CC_CODE='0'"
        If Not .IsEOF Then
            O2_rif_s = .GetValue("cc_value")
        End If
        
        'NR. MAX PARAMETRI PER PAGINA
        strSQL = "SELECT * FROM WAS_CONFIG WHERE cc_stationcode = '" & SelStation & "' and CC_CODE='80'"
        .selectionfast (strSQL)
        If Not .IsEOF Then
            PARAMETRI_PER_PAGINA = .GetValue("cc_value")
        End If
        If PARAMETRI_PER_PAGINA > 0 Then
            If MaxPar > PARAMETRI_PER_PAGINA Then
                Call windasLog("CreateReportInizializza: Troppi parametri per pagina!")
            End If
        Else
            Call windasLog("CreateReportInizializza: Numero massimo di parametri non configurato nel campo 80!")
        End If
        
        
        '********** estrae i codici di validità ***********
        .selectionfast "SELECT GT_Code FROM wds_GenTab WHERE GT_Type = 'VALID' AND (GT_NumInt <> 0 AND GT_NumInt IS NOT NULL) "
        Do While Not .IsEOF
          strValid = strValid & .ParSQLStr(.GetValue("GT_Code")) & ","
          .movenext
        Loop
        If (strValid <> "") Then
          strValid = Left(CStr(strValid), Len(CStr(strValid)) - 1)
        End If
        
        'Federica settembre 2017 - 10 minuti CO
        .selectionfast "SELECT CC_VALUE FROM was_config WHERE CC_StationCode = '" & SelStation & "' AND CC_CODE = '54'"
        If Not .IsEOF Then
            Abilita10minCO = IIf(.GetValue("CC_Value") = "True", True, False)
        End If
        
        Call windasLog("Avvio report: " + ModelFileName + " Data: " + Str(StartDate) + " Stazione: " + SelStation + " tabella: " + SelTable)
    
     End With
     Exit Sub

GestErrore:
     
     Call windasLog("CreateReportInizializza: " + Error(Err))

End Sub


Sub StatisticheDlgs133()

    Dim DayCount
    Dim CurRow
    Dim CurCol
    Dim RigaSemiora
    Dim anno
    Dim Valore
    Dim NumMinori
    Dim NumValidi
    Dim z
    Dim MezzoreNormFunz
    Dim MezzoreNormFunzDay
    Dim gg
    Dim ContaOreOK(7)
    Dim ContaOreNonOK(7)
    Dim Conta30MinAOver(7)
    Dim Conta30MinBok(7)
    Dim mm
    Dim ContaGiornoOK(7)
    Dim SommaValidiGiorno(7)
    Dim ContaGiornoNonOK(7)
    Dim MediaGiornoNonOK(7)
    Dim MediaGiornoOK(7)
    Dim SuperoGiorno(7)
    Dim GiornoNelLimite(7)
    Dim NumeroGiorniInMarcia
    Dim ndx
    Dim DataQuery As String

    On Error GoTo GestErrore
    
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    With rsExtractor
        
        'Ciclo per parametro
        For i = 0 To MaxPar
        
            '***** Scrittura del Nome parametro nell'intestazione *****
            objEngine.StringValue_CoordRange(0, 4, 2 + i) = CodPar(i, 1) & " " & CodPar(i, 2)
            
            '***** estrazione limite semiorario 100%(C38), limite semiorario 97% / 95% (C39) , limite giornaliero (C30) *****
            strSQL = "SELECT C38, C39, C30 FROM was_measures where C2 = '" & CodPar(i, 0) & "' AND CM_STATIONCODE = '" & SelStation & "'"
            .selectionfast CStr(strSQL)
            
            If (Not .IsEOF) Then
                '***** inserimento limite semiorario 100% *****
                If .GetValue("C38") > 0 Then
                    objEngine.NumberValue_CoordRange(0, 8, 2 + i) = .GetValue("C38")
                    LimiteSemiora100(i) = .GetValue("C38")
                Else
                    objEngine.StringValue_CoordRange(0, 8, 2 + i) = "n.a."
                End If
                
                '***** inserimento limite semiorario 97% *****
                If .GetValue("C39") > 0 Then
                    objEngine.NumberValue_CoordRange(0, 10, 2 + i) = .GetValue("C39")
                    LimiteSemiora97(i) = .GetValue("C39")
                Else
                    objEngine.StringValue_CoordRange(0, 10, 2 + i) = "n.a."
                End If
                
                '***** inserimento limite giornaliero *****
                If .GetValue("C30") > 0 Then
                    objEngine.NumberValue_CoordRange(0, 16, 2 + i) = .GetValue("C30")
                    LimiteGiorno(i) = .GetValue("C30")
                Else
                    objEngine.StringValue_CoordRange(0, 16, 2 + i) = "n.a."
                End If
            Else
                objEngine.StringValue_CoordRange(0, 8, 2 + i) = ""
                objEngine.StringValue_CoordRange(0, 10, 2 + i) = ""
                objEngine.StringValue_CoordRange(0, 16, 2 + i) = ""
            End If
        
        Next
       
        NumeroGiorniInMarcia = 0
        CheckDate = StartDate
        
        Do While CheckDate <= EndDate
             
            DataQuery = CreateDateForDB(CheckDate)
            MezzoreNormFunzDay = 0
            
            'Azzeramento matrice stato impianto
            For z = 0 To 48
                NFunzValue(z) = -9999
            Next
             
            'Azzeramento variabili
            For i = 0 To MaxPar
                ContaGiornoOK(i) = 0
                SommaValidiGiorno(i) = 0
                ContaGiornoNonOK(i) = 0
            Next

            'Estrazione medie semiorarie
            strSQL = "SELECT DT_MEASURECOD,DT_VALUE,DT_VALUE,DT_HOUR,DT_VALIDFLAG FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND DT_DATE = '" & DataQuery & "'"
            strSQL = strSQL & " ORDER BY DT_HOUR"
    
            .selectionfast CStr((strSQL))
        
            If Not .IsEOF Then
    
                 'controllo se impianto è in minimo tecnico
                 .m_Filter = "DT_MeasureCod = '" & NFunzPar & "'"
                 
                 Do While Not rsExtractor.IsEOF
                     NFunzValue(CLng(Left(.GetValue("DT_HOUR"), 2)) * 2 + CalcolaRigaSemiora(CLng(Right(.GetValue("DT_HOUR"), 2)))) = .GetValue("DT_VALUE")
                     .movenext
                 Loop
    
                 'inserimento stato impianto
                 For z = 0 To 47
                     If NFunzValue(z) <> -9999 Then
                         If NFunzValue(z) = 30 Then
                             MezzoreNormFunz = MezzoreNormFunz + 1
                             MezzoreNormFunzDay = MezzoreNormFunzDay + 1
                         End If
                     End If
                 Next
                 
                 If CLng(MezzoreNormFunzDay) >= 12 Then NumeroGiorniInMarcia = NumeroGiorniInMarcia + 1
    
                 '***** elaborazione dati per statistiche medie mezzora e giornaliere **************
                 For i = 0 To MaxPar
                 
                     .m_Filter = "DT_MeasureCod = '" & CodPar(i, 0) & "'"
                     
                     Do While Not rsExtractor.IsEOF

                         If NFunzValue(CLng(Left(.GetValue("DT_HOUR"), 2)) * 2 + CalcolaRigaSemiora(CLng(Right(.GetValue("DT_HOUR"), 2)))) <> -9999 Then
                             If NFunzValue(CLng(Left(.GetValue("DT_HOUR"), 2)) * 2 + CalcolaRigaSemiora(CLng(Right(.GetValue("DT_HOUR"), 2)))) = 30 Then
                                 
                                 If .GetValue("dt_validflag") = "VAL" Then
                                     
                                     'contatore di tutte le medie semiorarie valide
                                     ContaOreOK(i) = ContaOreOK(i) + 1
                                     
                                     'contatore e sommatoria per calcolo media giornaliera
                                     ContaGiornoOK(i) = ContaGiornoOK(i) + 1
                                     SommaValidiGiorno(i) = SommaValidiGiorno(i) + .GetValue("DT_VALUE")
                                     
                                     'controllo il numero di medie semiorarie valide < limite 100% (colonna A)
                                     If LimiteSemiora100(i) > 0 Then
                                        If .GetValue("DT_VALUE") > CLng(LimiteSemiora100(i)) Then
                                            Conta30MinAOver(i) = Conta30MinAOver(i) + 1
                                        End If
                                     End If
                                     
                                     'controllo il numero di medie semiorarie valide < limite 97% (colonna B)
                                     If LimiteSemiora97(i) > 0 Then
                                        If .GetValue("DT_VALUE") < CLng(LimiteSemiora97(i)) Then
                                            Conta30MinBok(i) = Conta30MinBok(i) + 1
                                        End If
                                     Else
                                        Conta30MinBok(i) = Conta30MinBok(i) + 1
                                     End If
                                     
                                 ElseIf .GetValue("dt_validflag") <> "VAL" Or .GetValue("dt_validflag") <> "AUX" Then
                                    'contatore di tutte le medie semiorarie non valide
                                    ContaOreNonOK(i) = ContaOreNonOK(i) + 1
                                    'contatore delle medie semiorarie non valide per il calcolo della media giornaliera
                                    ContaGiornoNonOK(i) = ContaGiornoNonOK(i) + 1
                                 End If
                                 
                             End If
                             
                         End If

                         .movenext
                         
                     Loop
    
                 Next
    
                 .m_Filter = ""
                 
            End If
            
            '***** calcoli per statistiche medie giornaliere *****
            For i = 0 To MaxPar
            
                If CLng(MezzoreNormFunzDay) > 11 Then
                    If CLng(ContaGiornoNonOK(i)) > 5 Then
                        'medie giornaliere non valide
                        MediaGiornoNonOK(i) = MediaGiornoNonOK(i) + 1
                    Else
                        If CLng(ContaGiornoOK(i)) > 0 Then
                            'medie giornaliere valide
                            MediaGiornoOK(i) = MediaGiornoOK(i) + 1
                            'numero medie giornaliere > limite
                            If LimiteGiorno(i) > 0 Then
                                If CLng(SommaValidiGiorno(i) / ContaGiornoOK(i)) > CLng(LimiteGiorno(i)) Then
                                    SuperoGiorno(i) = SuperoGiorno(i) + 1
                                Else
                                    GiornoNelLimite(i) = GiornoNelLimite(i) + 1
                                End If
                            Else
                                GiornoNelLimite(i) = GiornoNelLimite(i) + 1
                            End If
                        End If
                    End If
                End If
                
            Next
            
            gg = gg + 1
            CheckDate = DateAdd("d", 1, CheckDate)
            
        Loop

        '***** Scrittura dell'intestazione *****
        objEngine.StringValue_StringRange(0, "B2") = NomeImpianto
        objEngine.StringValue_StringRange(0, "B4") = "Report prodotto il " & FormatDateTime(Now, 0) & "   - Elaborazione conforme DL n. 133 del 11 Maggio 2005"
        objEngine.StringValue_StringRange(0, "B5") = "Dal " & FormatDateTime(StartDate, 0) & " al " & FormatDateTime(EndDate, 0) & " [Semiore di marcia impianto: " & CStr(MezzoreNormFunz) & "]"
        
        For i = 0 To MaxPar
            
            'numero di medie semiorarie valide (codice = "00")
            objEngine.NumberValue_CoordRange(0, 5, 2 + i) = Int(ContaOreOK(i))
            
            'numero di medie semiorarie non valide valide (codice = "15")
            objEngine.NumberValue_CoordRange(0, 6, 2 + i) = CLng(ContaOreNonOK(i))
            
            'percentuale di medie semiorarie valide
            If ContaOreOK(i) + ContaOreNonOK(i) > 0 Then
                objEngine.NumberValue_CoordRange(0, 7, 2 + i) = Int((ContaOreOK(i) / (ContaOreOK(i) + ContaOreNonOK(i))) * 100)
            Else
                objEngine.NumberValue_CoordRange(0, 7, 2 + i) = 0
            End If
            
            'numero di medie semiorarie superiori al limite 100% (colonna A)
            objEngine.NumberValue_CoordRange(0, 9, 2 + i) = CLng(Conta30MinAOver(i))

            'numero di medie semiorarie inferiori al limite 97% (colonna B)
            objEngine.NumberValue_CoordRange(0, 11, 2 + i) = Conta30MinBok(i)
                
            '% di medie semiorarie inferiori al limite 97% (colonna B)
            If ContaOreOK(i) > 0 Then
                objEngine.NumberValue_CoordRange(0, 12, 2 + i) = Int(Conta30MinBok(i) / ContaOreOK(i) * 100)
            Else
                objEngine.NumberValue_CoordRange(0, 12, 2 + i) = 0
            End If
            
            'numero di giornate in marcia (più di 12 ore di normale funzionamento)
            objEngine.NumberValue_CoordRange(0, 13, 2 + i) = Int(NumeroGiorniInMarcia)
        
            'numero di medie giornaliere valide
            objEngine.NumberValue_CoordRange(0, 14, 2 + i) = Int(MediaGiornoOK(i))
            
            'numero di medie giornaliere NON valide
            objEngine.NumberValue_CoordRange(0, 15, 2 + i) = Int(MediaGiornoNonOK(i))
            
            'numero di medie giornaliere > limite giornaliero
            objEngine.NumberValue_CoordRange(0, 17, 2 + i) = Int(SuperoGiorno(i))
            
            'numero medie giornaliere < limite giornaliero
            objEngine.NumberValue_CoordRange(0, 18, 2 + i) = Int(GiornoNelLimite(i))
        
        Next
        
    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("StatisticheDlgs133: " & Error(Err))

End Sub

Sub windasLog(evento As String)

    Dim ll As Integer
    
    On Error GoTo GestErrore
    
    If Dir("c:\Windas\log", vbDirectory) = "" Then
        MkDir "C:\Windas\log"
    End If
    If Dir("c:\Windas\log\logWindasReport", vbDirectory) = "" Then
        MkDir "c:\Windas\log\logWindasReport"
    End If

    ll = FreeFile
    Open "c:\Windas\log\logWindasReport" + "\" + Format(Now, "yyyymmdd") + ".txt" For Append As #ll
    Print #ll, Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + evento
    Close (ll)
    
    Debug.Print Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + evento
    
    Exit Sub
    
GestErrore:
    Debug.Print Error(Err)
    Resume fine
fine:
    
End Sub

Public Sub GiornalieroFM(AncheTransitorio As Boolean, Optional ByVal isSemiore As Boolean = False)
    
    Dim StatoImpianto As String
    Dim Min, Max As Double
    Dim nn As Integer
    Dim MediaDaVerificare(72) As Boolean
    Dim indiceOra As Integer
    Dim Conc_Value(48)
    
    On Error GoTo GestErrore
    
    StatoImpianto = "DT_CUSTOM1"
    
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = IIf(isSemiore, NumMedieSemiorario, NumMedieOrario) + 1
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    If isSemiore Then
        RIGA_LIMITI_PERIODO = RIGA_LIMITI_PERIODO + 3
        RIGA_TOTALE_PERIODO = RIGA_TOTALE_PERIODO + 3
    
        Call ScriviPartiFisse("RPT_GIORNO", "COL_IMPIANTO", "TIPO_STATOIMP", "ROW_LIMITI_SEMIORA", "ROW_TOT_GIORNO", _
            "ROW_LIMITI_GIORNO", Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy"), "COL_ORA")
    Else
        Call ScriviPartiFisse("RPT_GIORNO", "COL_IMPIANTO", "TIPO_STATOIMP", "ROW_LIMITI_ORA", "ROW_TOT_GIORNO", _
            "ROW_LIMITI_GIORNO", Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy"), "COL_ORA")
    End If
    Call LegendaNote("GF")
    For z = RIGA_LEGENDA_NOTE + 2 To RIGA_LEGENDA_NOTE + 5
        objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc(IIf(isSemiore, "NOTE_OS_F_", "NOTE_O_F_") & CStr(z))
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** Azzeramento matrici *****
    nn = 0
   
    For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
        NFunzValue(z) = -9999
        MediaDaVerificare(i) = False
    Next
        
    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
        
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        'Federica luglio 2017 - gestione della traduzione del nome parametro in LocTable.ini
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" & CodPar(i, 2) & ")"
        If LimiteOraFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteOraFlussi(i)
        If LimiteGiornoFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_PERIODO, CurCol + POS_COL_VALORE) = LimiteGiornoFlussi(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    '************** estrae stato impianto **************
    strSQL = "SELECT * FROM " & SelTable & " WHERE " & _
            " DT_STATIONCODE = '" & SelStation & "' AND DT_DATE = '" & String_StartDate_DB & "' AND DT_MeasureCod = '" & NFunzPar & "'"
    strSQL = strSQL & " ORDER BY DT_HOUR"
    With rsExtractor
        .selectionfast (strSQL)
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not rsExtractor.IsEOF
                indiceOra = Left(.GetValue("DT_HOUR"), 2)
                If isSemiore Then indiceOra = indiceOra * 2 + CalcolaRigaSemiora(Right(.GetValue("DT_HOUR"), 2))
                NFunzValue(indiceOra) = .GetValue(StatoImpianto)
                .movenext
            Loop
        End If
            
        'luca gennaio 2015 inserisco gestione tutti stati impianto (ABS UDINE)
        'inserimento stato impianto
        OreNormFunz = 0
        
        For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
            
            'Nicolò Agosto 2015 rendo più flessibile la gestione dei testi per lo stato impianto
            If CDbl(NFunzValue(z)) <> -9999 Then
                'If CDbl(NFunzValue(z)) = 30 Then OreNormFunz = OreNormFunz + 1
                CurRow = RIGA_INTESTAZIONI + 1 + z
                objEngine.StringValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = GetStsImpDescription(CInt(NFunzValue(z)))
                
                'luca 05/10/2016 gestisco marcia / marcia + transitorio
                If AncheTransitorio Then
                    If CDbl(NFunzValue(z)) = 30 Or CDbl(NFunzValue(z)) = 31 Or CDbl(NFunzValue(z)) = 32 Then
                        OreNormFunz = OreNormFunz + 1
                        MediaDaVerificare(z) = True
                    End If
                Else
                    If CDbl(NFunzValue(z)) = 30 Then
                        OreNormFunz = OreNormFunz + 1
                        MediaDaVerificare(z) = True
                    End If
                End If
                
            End If
        Next
            
        'luca gennaio 2015 scrivo le ore di normal funzionamento
        objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = OreNormFunz
    
        '************ estrae le concentrazioni e calcola il flusso di massa *************
        For i = 0 To MaxPar
        
            '****** Azzeramento matrici ******
            For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
                Conc_Value(z) = -9999
            Next
            Min = 999999
            Max = -999999
    
            strSQL = "SELECT * FROM " & SelTable & " WHERE " & _
                    " DT_STATIONCODE = '" & SelStation & "' AND DT_DATE = '" & String_StartDate_DB & "' AND DT_MeasureCod = '" & CodPar(i, 0) & "'"
            strSQL = strSQL & " ORDER BY DT_HOUR"
            .selectionfast (strSQL)
            
            If .RecordCount > 0 Then
              .MoveFirst
              Do While Not rsExtractor.IsEOF
                indiceOra = Left(.GetValue("DT_HOUR"), 2)
                If isSemiore Then indiceOra = indiceOra * 2 + CalcolaRigaSemiora(Right(.GetValue("DT_HOUR"), 2))
                Conc_Value(indiceOra) = .GetValue("DT_FM") 'Federica agosto 2017 - uso il dato calcolato
                .movenext
              Loop
            End If
            
            '***** inserisce le misure in tabella *****
            For z = 0 To IIf(isSemiore, NumMedieSemiorario, NumMedieOrario)
                '*** se siamo in normale funzionamento..
                'luca 05/10/2016
                'If CDbl(NFunzValue(z)) = 30 Then.
                CurRow = RIGA_INTESTAZIONI + 1 + z
                CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
                If MediaDaVerificare(z) Then
                    If Conc_Value(z) = -9999 Then
                        'luca 25/01/2016 valid flag piemonte
                        objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "AL", "'(4)")
                     Else
                        objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = Conc_Value(z)
                        
                        'luca 19/01/2016 controllo eventuale supero
                        If LimiteOraFlussi(i) > 0 Then
                            If Conc_Value(z) > LimiteOraFlussi(i) Then
                                'luca 25/01/2016 valid flag piemonte
                                objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = IIf(IsPiemonte, "S", "(5)")
                            End If
                        End If
                        
                        If Conc_Value(z) < Min Then Min = Conc_Value(z)
                        If Conc_Value(z) > Max Then Max = Conc_Value(z)
                    End If
                End If
            Next
            
            '**** minimo/massimo
            If Min <> 999999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = FormatNumber(Min, 2)
            If Max <> -999999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = FormatNumber(Max, 2)
        
            '***** Riepilogo finale per parametro *****
            Call RiepilogoParametroFlussi("wds_days", i, String_StartDate_DB, String_EndDate_DB, LimiteGiornoFlussi(i))
            objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
            objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
        Next
    End With
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    Set rsValMin = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("GiornalieroFM: " & Error(Err))

End Sub

Public Sub Mensile(numMedie As Integer, minimoMedieGiorno As Integer, minimoMedieMese As Integer, isSemiore As Boolean)
    
    Dim nn As Integer 'vale sempre 0....
    
    On Error GoTo GestErrore
   
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 31
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    If isSemiore Then
        Call ScriviPartiFisse("RPT_MESE", "COL_SEMIORE_IMPIANTO", "", "ROW_LIMITI_GIORNO", "ROW_MEDIA_MESE", "ROW_LIMITI_MESE", _
            mese(month(StartDate)) & " " & CStr(year(StartDate)), "COL_GIORNI")
    Else
        Call ScriviPartiFisse("RPT_MESE", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_GIORNO", "ROW_MEDIA_MESE", "ROW_LIMITI_MESE", _
            mese(month(StartDate)) & " " & CStr(year(StartDate)), "COL_GIORNI")
    End If
    Call LegendaNote("M")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 11
        If isSemiore Then
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_MS_" & CStr(z))
        Else
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_M_" & CStr(z))
        End If
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" + CodPar(i, 2) + ")"
        If LimiteGiorno(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteGiorno(i)
        If LimiteMese(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_PERIODO, CurCol + POS_COL_VALORE) = LimiteMese(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** Stato impianto *****
    TotOreNormFunz = 0
    strSQL = "SELECT DT_DATE, DT_NRTOT FROM wds_days " & _
             "WHERE DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' " & _
             "AND DT_STATIONCODE = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
            CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 7, 2))
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")

            .movenext
        Loop
    End With
    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz

    '***** MEDIE *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        
        With rsExtractor
            strSQL = "SELECT * FROM wds_days " & _
                     "WHERE DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' " & _
                     "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                     "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                     "ORDER BY DT_DATE"
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                    CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 7, 2))
                    
                    '***** Media giornaliera *****
                    Call ValidaMedia(minimoMedieGiorno, LimiteGiorno(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametro("wds_month", i, Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), MinimoOreMese, LimiteMese(i))
                Call MinimoMassimo("wds_days", String_StartDate_DB, String_EndDate_DB, CodPar(i, 0))
                'luca marzo 2018
                'objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("Mensile: " & Error(Err))
    Resume Next

End Sub

Function mese(Nmese)
    
    Select Case Nmese
        Case 1:
            mese = Loc("GENNAIO")
        
        Case 2:
            mese = Loc("FEBBRAIO")
        
        Case 3:
            mese = Loc("MARZO")
        
        Case 4:
            mese = Loc("APRILE")
        
        Case 5:
            mese = Loc("MAGGIO")
        
        Case 6:
            mese = Loc("GIUGNO")
            
        Case 7:
            mese = Loc("LUGLIO")
        
        Case 8:
            mese = Loc("AGOSTO")
        
        Case 9:
            mese = Loc("SETTEMBRE")
        
        Case 10:
            mese = Loc("OTTOBRE")
        
        Case 11:
            mese = Loc("NOVEMBRE")
            
        Case 12:
            mese = Loc("DICEMBRE")
    End Select

End Function



Public Sub MensileFM(Optional ByVal isSemiore As Boolean = False)
    
    Dim nn As Integer
    
    On Error GoTo GestErrore
    
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 31
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    If isSemiore Then
        Call ScriviPartiFisse("RPT_MESE_FM", "COL_SEMIORE_IMPIANTO", "", "ROW_LIMITI_GIORNO", "ROW_TOT_MESE", "ROW_LIMITI_MESE", _
            mese(month(StartDate)) & " " & CStr(year(StartDate)), "COL_GIORNI")
    Else
        Call ScriviPartiFisse("RPT_MESE_FM", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_GIORNO", "ROW_TOT_MESE", "ROW_LIMITI_MESE", _
            mese(month(StartDate)) & " " & CStr(year(StartDate)), "COL_GIORNI")
    End If
    Call LegendaNote("MF")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 10
        objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_M_F_" & CStr(z))
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" & CodPar(i, 3) & ")"
        If LimiteGiornoFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteGiornoFlussi(i)
        If LimiteMeseFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_PERIODO, CurCol + POS_COL_VALORE) = LimiteMeseFlussi(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** Stato impianto *****
    strSQL = "SELECT DT_DATE, DT_NRTOT FROM wds_days " & _
             "WHERE DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' " & _
             "AND DT_STATIONCODE = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
            CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 7, 2))
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")

            .movenext
        Loop
    End With
    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz
    
    '***** FLUSSI *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        
        With rsExtractor
            strSQL = "SELECT * FROM wds_days " & _
                     "WHERE DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' " & _
                     "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                     "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                     "ORDER BY DT_DATE"
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                    CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 7, 2))
                    
                    '***** Flusso giornaliero *****
                    Call ValidaFlusso(LimiteGiornoFlussi(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametroFlussi("wds_month", i, Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), LimiteMeseFlussi(i))
                Call MinimoMassimoFlussi("wds_days", String_StartDate_DB, String_EndDate_DB, CodPar(i, 0))
                objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                'objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub

GestErrore:
    Call windasLog("MensileFM: " & Error(Err))
    
End Sub


Public Sub Annuale(numMedie As Integer, minimoMedieMese As Integer, isSemiore As Boolean)
    
    On Error GoTo GestErrore

    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 12 'Mesi
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    If isSemiore Then
        Call ScriviPartiFisse("RPT_ANNO", "COL_SEMIORE_IMPIANTO", "", "ROW_LIMITI_MESE", "ROW_MEDIA_ANNO", "", Loc("ANNO") & ": " & year(StartDate), "COL_MESI")
    Else
        Call ScriviPartiFisse("RPT_ANNO", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_MESE", "ROW_MEDIA_ANNO", "", Loc("ANNO") & ": " & year(StartDate), "COL_MESI")
    End If
    Call LegendaNote("A")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 10
        If isSemiore Then
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_AS_" & CStr(z))
        Else
            objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_A_" & CStr(z))
        End If
    Next z
    'Mesi
    For z = 1 To 12
        objEngine.StringValue_StringRange(0, "B" & CStr(z + 7)) = mese(z)
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
          
    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" + CodPar(i, 2) + ")"
        If LimiteMese(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteMese(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    '***** Stato impianto *****
    TotOreNormFunz = 0
    strSQL = "SELECT DT_DATE, DT_NRTOT FROM wds_month " & _
             "WHERE DT_DATE BETWEEN '" & Left(String_StartDate_DB, 6) & "' AND '" & Left(String_EndDate_DB, 6) & "' " & _
             "AND DT_STATIONCODE = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
            CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 5, 2))
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")

            .movenext
        Loop
    End With
    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz
    
    '***** MEDIE *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        
        strSQL = "SELECT * FROM wds_month " & _
                 "WHERE DT_DATE BETWEEN '" & Left(String_StartDate_DB, 6) & "' AND '" & Left(String_EndDate_DB, 6) & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                 "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                 "ORDER BY DT_DATE"
        With rsExtractor
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                    CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 5, 2))
                    
                    '***** Media mensile *****
                    Call ValidaMedia(minimoMedieMese, LimiteMese(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametro("wds_year", i, Left(String_StartDate_DB, 4), Left(String_EndDate_DB, 4), 0, LimiteAnno(i))
                Call MinimoMassimo("wds_month", Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), CodPar(i, 0))
                If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("Annuale: " & Error(Err))
    Resume Next
End Sub

Public Sub AnnualeFM(isSemiore As Boolean)
    
    Dim nn As Integer
    
    On Error GoTo GestErrore
    
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 12
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    If isSemiore Then
        Call ScriviPartiFisse("RPT_ANNO_FM", "COL_SEMIORE_IMPIANTO", "", "ROW_LIMITI_MESE", "ROW_TOT_ANNO", "", Loc("ANNO") & year(StartDate), "COL_MESI")
    Else
        Call ScriviPartiFisse("RPT_ANNO_FM", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_MESE", "ROW_TOT_ANNO", "", Loc("ANNO") & year(StartDate), "COL_MESI")
    End If
    Call LegendaNote("AF")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 9
        objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc(IIf(isSemiore, "NOTE_AS_F_", "NOTE_A_F_") & CStr(z))
    Next z
    'Mesi
    For z = 1 To 12
        objEngine.StringValue_StringRange(0, "B" & CStr(z + 7)) = mese(z)
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" & CodPar(i, 3) & ")"
        If LimiteMeseFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteMeseFlussi(i)
    Next i
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    '***** Stato impianto *****
    strSQL = "SELECT DT_DATE, DT_NRTOT FROM wds_month " & _
             "WHERE DT_DATE BETWEEN '" & Left(String_StartDate_DB, 6) & "' AND '" & Left(String_EndDate_DB, 6) & "' " & _
             "AND DT_STATIONCODE = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
            CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 5, 2))
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")

            .movenext
        Loop
    End With
    objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz
    
    '***** FLUSI *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
    
        strSQL = "SELECT * FROM wds_month " & _
                 "WHERE DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                 "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                 "ORDER BY DT_DATE"
        With rsExtractor
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                    CurRow = RIGA_INTESTAZIONI + CInt(Mid(.GetValue("DT_DATE"), 5, 2))
                    
                    '***** Flusso mensile *****
                    Call ValidaFlusso(LimiteMeseFlussi(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametroFlussi("wds_month", i, Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), LimiteAnnoFlussi(i))
                Call MinimoMassimoFlussi("wds_month", Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), CodPar(i, 0))
                If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                'objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("AnnualeFM: " & Error(Err))
    Resume Next
End Sub

'Federica luglio 2017 - Estrazione di TUTTI i limiti dei parametri
Private Sub GetLimitiParametri()
    
    On Error GoTo GestErrore
        
    Dim indice As Integer

    For indice = 0 To MaxPar
        'Azzero gli array
        LimiteOra(indice) = -9999
        LimiteGiorno(indice) = -9999
        LimiteMese(indice) = -9999
        LimiteAnno(indice) = -9999
        Limite48H(indice) = -9999
        LimiteTrimestre(indice) = -9999
        LimiteOraFlussi(indice) = -9999
        LimiteMeseFlussi(indice) = -9999
        LimiteGiornoFlussi(indice) = -9999
        LimiteAnnoFlussi(indice) = -9999
        LimiteTrimestreFlussi(indice) = -9999
        LimiteSemiora100(indice) = -9999
        LimiteSemiora97(indice) = -9999
                
        strSQL = "SELECT C28, C29, C30, C31, C32, C33, C34, C35, C54, C57, C64, C78, C79 " & _
                 "FROM was_measures where C2 = '" & CodPar(indice, 0) & "' and cm_stationcode = '" & SelStation & "'"
        With rsExtractor
            If (.selectionfast(strSQL)) Then
                LimiteSemiora100(indice) = Val(Replace(.GetValue("C28"), ",", "."))
                LimiteSemiora97(indice) = Val(Replace(.GetValue("C29"), ",", "."))
                LimiteOra(indice) = Val(Replace(.GetValue("C30"), ",", "."))
                LimiteGiorno(indice) = Val(Replace(.GetValue("C31"), ",", "."))
                LimiteMese(indice) = Val(Replace(.GetValue("C33"), ",", "."))
                LimiteAnno(indice) = Val(Replace(.GetValue("C57"), ",", "."))
                Limite48H(indice) = Val(Replace(.GetValue("C32"), ",", "."))
                LimiteTrimestre(indice) = Val(Replace(.GetValue("C78"), ",", "."))
                LimiteOraFlussi(indice) = Val(Replace(.GetValue("C64"), ",", "."))
                LimiteMeseFlussi(indice) = Val(Replace(.GetValue("C34"), ",", "."))
                LimiteGiornoFlussi(indice) = Val(Replace(.GetValue("C54"), ",", "."))
                LimiteAnnoFlussi(indice) = Val(Replace(.GetValue("C35"), ",", "."))
                LimiteTrimestreFlussi(indice) = Val(Replace(.GetValue("C79"), ",", "."))
            End If
        End With
    
    Next indice
    
    Exit Sub
    
GestErrore:
    Call windasLog("GetLimitiParametri: " & Error(Err))

End Sub

Private Sub RiepilogoParametro(ByVal tabella As String, ByVal indice_parametro As String, ByVal dal As String, ByVal al As String, _
    ByVal MinimoOre As Integer, ByVal limite As Double)
    
    media_ = 0
    note_ = ""
    id_ = 0

    On Error GoTo GestErrore

    With rsExtractor
        strSQL = "SELECT DT_NRTOT, DT_NR, DT_VALUE, DT_VALIDFLAG FROM " & tabella & " WHERE DT_DATE between '" & dal & "' AND '" & al & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' and DT_MEASURECOD = '" & CodPar(indice_parametro, 0) & "' " 'AND DT_VALIDFLAG in ('VAL', 'VAH')"
        If (.selectionfast(strSQL)) Then
            media_ = .GetValue("DT_VALUE")
            If IsPiemonte Then note_ = "V"
            If media_ <> -9999 Then
                id_ = .GetValue("DT_NR") / .GetValue("DT_NRTOT") * 100
                If .GetValue("DT_NRTOT") < MinimoOre Then
                    note_ = IIf(IsPiemonte, "NS", "'(2)")
                    id_ = 0
                Else
                    If InStr("VAL, VAH", .GetValue("DT_VALIDFLAG")) = 0 Then
                        note_ = IIf(IsPiemonte, "NV", "'(4)")
                        id_ = 0
                    Else
                        If limite > 0 Then
                            If media_ > limite Then note_ = IIf(IsPiemonte, "S", "'(6)")
                        End If
                    End If
                End If
            Else
                'Gestione per report non in italiano
                note_ = IIf(Loc("LINGUA") = "IT", IIf(IsPiemonte, "NV", "'(2)"), "'(2)")
            End If
        End If
    End With
    
    Exit Sub
    
GestErrore:
    Call windasLog("RiepilogoParametro: " & Err())

End Sub

Private Sub RiepilogoParametroFlussi(ByVal tabella As String, ByVal indice_parametro As String, ByVal dal As String, ByVal al As String, _
    ByVal limite As Double)

    media_ = 0
    note_ = ""

    On Error GoTo GestErrore

    With rsExtractor
        strSQL = "SELECT COUNT(DT_FM) AS Tot, Sum(DT_FM) AS DT_FM FROM " & tabella & " WHERE DT_DATE between '" & dal & "' AND '" & al & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' and DT_MEASURECOD = '" & CodPar(indice_parametro, 0) & "' " & _
                 "AND DT_FM <> -9999"
        If (.selectionfast(strSQL)) Then
            If .GetValue("Tot") > 0 Then
                media_ = .GetValue("DT_FM")
                If limite > 0 Then
                    If media_ > limite Then note_ = IIf(IsPiemonte, "S", "'(6)")
                End If
            Else
                'Gestione per report non in italiano
                note_ = IIf(Loc("LINGUA") = "IT", IIf(IsPiemonte, "", "'(1)"), "'(4)")
            End If
        End If
    End With
    
    Exit Sub
    
GestErrore:
    Call windasLog("RiepilogoParametroFlussi: " & Err())

End Sub

Private Sub ValidaFlusso(ByVal limite As Double)

    media_ = 0
    note_ = ""
    id_ = 0
    
    On Error GoTo GestErrore
    
    Dim idx As Integer

    With rsExtractor
        '***** Flusso *****
        media_ = .GetValue("DT_FM")
        If IsPiemonte Then note_ = "V"
        
        '***** Se ho il flusso  *****
        If media_ <> -9999 Then
            If limite > 0 Then
                If media_ > limite Then
                    '***** Supero limite *****
                    note_ = IIf(IsPiemonte, "S", "'(5)")
                End If
            End If
        Else
            '***** Flusso non valido *****
            note_ = IIf(IsPiemonte, "NV", "'(4)")
        End If
    End With
    
    Exit Sub
    
GestErrore:
    Call windasLog("ValidaFlusso: " & Err())

End Sub

Private Sub ValidaMedia(ByVal minimo_ore As Integer, ByVal limite As Double)

    media_ = 0
    note_ = ""
    id_ = 0

    On Error GoTo GestErrore

    With rsExtractor
        '***** Media *****
        media_ = .GetValue("DT_VALUE")
        If media_ <> -9999 Then
            '***** Se ci sono almeno <minimo_ore> ore di normal funzionamento *****
            If .GetValue("DT_NRTOT") >= minimo_ore Then
                '***** Status di validità media *****
                If InStr(strValid, .GetValue("DT_VALIDFLAG")) Then
                    If IsPiemonte Then note_ = "V"
                    If limite > 0 Then
                        If .GetValue("DT_VALUE") > limite Then
                            '***** Supero limite *****
                            note_ = IIf(IsPiemonte, "S", "'(5)")
                        End If
                    End If
                    
                Else
                    '***** Media non valida *****
                    note_ = IIf(IsPiemonte, "NV", "'(4)")
                End If
                
                '***** ID% *****
                id_ = .GetValue("DT_NR") / .GetValue("DT_NRTOT") * 100
            Else
                '***** Media non significativa *****
                note_ = IIf(Loc("LINGUA") = "IT", IIf(IsPiemonte, "NS", "'(2)"), "'(4)")
            End If
        Else
            'Gestita per report non in italiano
            note_ = IIf(Loc("LINGUA") = "IT", IIf(IsPiemonte, "NS", "'(2)"), "'(4)")
        End If
    End With

    Exit Sub
    
GestErrore:
    Call windasLog("ValidaMedia: " & Err())

End Sub

Public Sub Trimestrale(minimoMedieMese As Integer, isSemiore As Boolean)
    
    'Federica luglio 2017
    Dim daData As Date
    Dim aData As Date
    Dim nn As Integer
    Dim i As Integer
    Dim pos As Integer
    
    On Error GoTo GestErrore

    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 4
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    Call ScriviPartiFisse("RPT_TRIMESTRE", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_TRIMESTRE", "ROW_MEDIA_ANNO", "", Loc("ANNO") & year(StartDate), "COL_TRIMESTRI")
    Call LegendaNote("T")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 10
        objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_T_" & CStr(z))
    Next z
    'Trimestri
    pos = 1
    For z = 1 To 12 Step 3
        objEngine.StringValue_StringRange(0, "B" & CStr(pos + offset_riga)) = mese(z) & " " & mese(z + 2)
        pos = pos + 1
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    Call GetLimitiParametri   'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (" + CodPar(i, 2) + ")"
        If LimiteTrimestre(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteTrimestre(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** STATO IMPIANTO *****
    TotOreNormFunz = 0
    'Estraggo i dati per i 4 trimestri
    strSQL = "SELECT * FROM WDS_QUARTERLY " & _
             "WHERE LEFT(DT_DATE, 4) = '" & CStr(year(StartDate)) & "' " & _
             "AND DT_STATIONCODE  = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
            'Calcolo la riga del trimestre
            Select Case Mid(.GetValue("DT_DATE"), 5, 2)
                Case "01"
                    CurRow = offset_riga + 0
                Case "04"
                    CurRow = offset_riga + 1
                Case "07"
                    CurRow = offset_riga + 2
                Case "09"
                    CurRow = offset_riga + 3
            End Select
                            
            '****** Inserimento numero ore di normale funzionamento ******
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            'Aggiorno contatore Totale delle ore normal funzionamento
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")
            
            .movenext
        Loop
        
        '***** Totale ore di normal funzionamento *****
        objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz
    End With
    
    '***** MEDIE *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        
        'Estraggo i dati per i 4 trimestri
        strSQL = "SELECT * FROM WDS_QUARTERLY WHERE LEFT(DT_DATE, 4) = '" & CStr(year(StartDate)) & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                 "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                 "ORDER BY DT_DATE"
        With rsExtractor
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                    'Calcolo la riga del trimestre
                    Select Case Mid(.GetValue("DT_DATE"), 5, 2)
                        Case "01"
                            CurRow = offset_riga + 0
                        Case "04"
                            CurRow = offset_riga + 1
                        Case "07"
                            CurRow = offset_riga + 2
                        Case "09"
                            CurRow = offset_riga + 3
                    End Select
                                    
                    '***** Media trimestrale *****
                    Call ValidaMedia(minimoMedieMese, LimiteTrimestre(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametro("wds_quarterly", i, Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), MinimoOreMese, LimiteTrimestre(i))
                Call MinimoMassimo("wds_quarterly", Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), CodPar(i, 0))
                If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                If id_ > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("Annuale: " & Error(Err))
    Resume Next
End Sub


Public Sub TrimestraleFM(isSemiore As Boolean)
    
    'Federica luglio 2017
    Dim daData As Date
    Dim aData As Date
    Dim nn As Integer
    Dim i As Integer
    Dim pos As Integer
    
    On Error GoTo GestErrore
    
    'Federica agosto 2017 - Standardizzato inizio report >>>>>
    '***** Gestione posizionamenti nel foglio *****
    NR_RIGHE = 4
    Call CalcolaPosizioniFoglio
    Call IniziaReport
    Call ScriviPartiFisse("RPT_TRIMESTRE_FM", "COL_ORE_IMPIANTO", "", "ROW_LIMITI_TRIMESTRE", "ROW_TOT_ANNO", "", Loc("ANNO") & year(StartDate), "COL_TRIMESTRI")
    Call LegendaNote("T")
    For z = RIGA_LIMITI_PERIODO + 7 To RIGA_LIMITI_PERIODO + 10
        objEngine.StringValue_StringRange(0, "D" & CStr(z)) = Loc("NOTE_T_" & CStr(z))
    Next z
    'Trimestri
    pos = 1
    For z = 1 To 12 Step 3
        objEngine.StringValue_StringRange(0, "B" & CStr(pos + offset_riga)) = mese(z) & " " & mese(z + 2)
        pos = pos + 1
    Next z
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Call GetLimitiParametri 'Federica agosto 2017 - Estrazione di TUTTI i limiti
    
    'Federica agosto 2017 - Semplificata scrittura dei dati parametro >>>
    '***** Scrittura dati e limiti per parametro *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        objEngine.StringValue_CoordRange(0, RIGA_PARAMETRI, CurCol + POS_COL_NOTE) = IIf(Loc(CodPar(i, 1)) = "---", CodPar(i, 1), Loc(CodPar(i, 1))) + " (Kg)"
        If LimiteTrimestreFlussi(i) > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_LIMITI_SINGOLI, CurCol + POS_COL_VALORE) = LimiteTrimestreFlussi(i)
    Next
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '***** STATO IMPIANTO *****
    'Estraggo i dati per i 4 trimestri
    strSQL = "SELECT * FROM WDS_QUARTERLY WHERE LEFT(DT_DATE, 4) = '" & CStr(year(StartDate)) & "' " & _
             "AND DT_STATIONCODE  = '" & SelStation & "' " & _
             "AND DT_MEASURECOD = '" & NFunzPar & "' " & _
             "ORDER BY DT_DATE"
    With rsExtractor
        .selectionfast (strSQL)
        Do While Not .IsEOF
        
            Select Case Mid(.GetValue("DT_DATE"), 5, 2)
                Case "01"
                    CurRow = offset_riga + 0
                Case "04"
                    CurRow = offset_riga + 1
                Case "07"
                    CurRow = offset_riga + 2
                Case "09"
                    CurRow = offset_riga + 3
            End Select
                            
            '****** Inserimento numero ore di normale funzionamento ******
            objEngine.NumberValue_CoordRange(0, CurRow, COLONNA_IMPIANTO) = .GetValue("DT_NRTOT")
            'Aggiorno contatore Totale delle ore normal funzionamento
            TotOreNormFunz = TotOreNormFunz + .GetValue("DT_NRTOT")
            
            .movenext
        Loop
        
        '***** Totale ore di normal funzionamento *****
        objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, COLONNA_IMPIANTO) = TotOreNormFunz
    End With
    
    '***** FUSSI *****
    For i = 0 To MaxPar
        CurCol = COLONNA_PARAMETRI + (COLONNE_PER_PARAMETRO * i)
        
        'Estraggo i dati per i 4 trimestri
        strSQL = "SELECT * FROM WDS_QUARTERLY WHERE LEFT(DT_DATE, 4) = '" & CStr(year(StartDate)) & "' " & _
                 "AND DT_STATIONCODE  = '" & SelStation & "' " & _
                 "AND DT_MEASURECOD = '" & CodPar(i, 0) & "' " & _
                 "ORDER BY DT_DATE"
        With rsExtractor
            If (.selectionfast(strSQL)) Then
                Do While Not .IsEOF
                
                    Select Case Mid(.GetValue("DT_DATE"), 5, 2)
                        Case "01"
                            CurRow = offset_riga + 0
                        Case "04"
                            CurRow = offset_riga + 1
                        Case "07"
                            CurRow = offset_riga + 2
                        Case "09"
                            CurRow = offset_riga + 3
                    End Select
                                    
                    '***** Media trimestrale *****
                    Call ValidaFlusso(LimiteTrimestreFlussi(i))
                    If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_VALORE) = media_
                    If note_ <> "" Then objEngine.StringValue_CoordRange(0, CurRow, CurCol + POS_COL_NOTE) = note_
                    If id_ > 0 Then objEngine.NumberValue_CoordRange(0, CurRow, CurCol + POS_COL_ID) = id_
                    
                    .movenext
                Loop
            
                '***** Riepilogo finale per parametro *****
                Call RiepilogoParametroFlussi("wds_quarterly", i, Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), LimiteTrimestreFlussi(i))
                Call MinimoMassimoFlussi("wds_quarterly", Left(String_StartDate_DB, 6), Left(String_EndDate_DB, 6), CodPar(i, 0))
                If media_ <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_VALORE) = media_
                If note_ <> "" Then objEngine.StringValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_NOTE) = note_
                'If id_ > 0 Then objEngine.NumberValue_CoordRange(0, RIGA_TOTALE_PERIODO, CurCol + POS_COL_ID) = id_
                If xMin <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MINIMO, CurCol + POS_COL_VALORE) = xMin
                If xMax <> -9999 Then objEngine.NumberValue_CoordRange(0, RIGA_MASSIMO, CurCol + POS_COL_VALORE) = xMax
            End If
        End With
    Next i
    
    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("Annuale: " & Error(Err))
    Resume Next
End Sub
Public Sub QAL2()
    
    Dim ParDataQAL2 As String
    Dim ParCampoValore As String
    Dim ParCampoValidita As String
    Dim ParCampoStsImp As String
    Dim TipoStatoImpianto  As String
    Dim RangePar  As String
    Dim valid_flag As String
    Dim ParRangeTaratura As String
    Dim ValoreRifStatoImpianto As Integer
    Dim TipoMedia As Integer
    Dim MezzoreNormFunz As Integer
    Dim misura As Integer
    Dim StartWeek As Date
    Dim EndWeek As Date
    Dim ContaOreOKSetteGiorni(99) As Double
    Dim ContatoreNonOKTarSetteGiorni(99) As Double
    Dim CinquePerCento(99) As Double
    Dim QuarantaPerCento(99) As Double
    Dim RangeTar(99) As Double
    Dim codLinea As Integer
    Dim sDataInizioParametro As String
    
    Const DOMENICA = 1
    Const LUNEDI = 2
    
    On Error GoTo GestErrore
    
    xStartRange = 6
    yStartRange = 3
    xEndRange = 10
    yEndRange = 7
    yDestTopLeft = 3
    
    Const MEDIA_ORARIA = 0
    Const MEDIA_SEMIORARIA = 1
    Const NF_DA_PERCENTUALE = 0       'assegna Norm. Funz. in base al valore di NFunzPar espresso in percentuale
    Const NF_DA_VALORE = 1            'assegna Norm. Funz. in base al valore di NFunzPar
    
    
    codLinea = 0 'Right(ThisWorkbook.Worksheets("Settings").range("B12").Value, 1)
    numLinea = CInt(codLinea)
    
    '******* assegnazioni da parte dell'utilizzatore *******
    ParRangeTaratura = "C43"          'indica il campo in WAS_MEASURES utilizzato per salvare il Range di taratura
    ParDataQAL2 = "C45"               'indica il campo in WAS_MEASURES utilizzato per salvare la data di esecuzione della QAL2
    ParCampoValore = "DT_VALUE"       'Indica il campo in WDS_ELAB/WDS_HALF utilizzato per salvare la media elaborata
    ParCampoValidita = "DT_VALIDFLAG" 'Indica il campo in WDS_ELAB/WDS_HALF utilizzato per salvare il flag di validità della media elaborata
    ParCampoStsImp = "DT_CUSTOM1"       'Indica il campo in WDS_ELAB/WDS_HALF utilizzato per salvare lo stato impianto
    Const FORMATO_ORA_IN = "00.00"    'indica il formato della prima media oraria del giorno; alternativa: "00.30"/ "01.00"; usato nel calcolo dell'indice orario

    TipoStatoImpianto = NF_DA_VALORE  'indica che l'impianto è a regime se NFunzPar = ValoreRifStatoImpianto
    ValoreRifStatoImpianto = 30
    
    '************************************************************
    '* INIZIALIZZAZIONI                                         *
    '************************************************************
    
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    With rsExtractor
    
    
    'Nicola 05/03/2015 Scrivo Nome impianto
    objEngine.StringValue_StringRange(0, "A2") = NomeImpianto
    
    '***** Scrittura della data nell'intestazione *****
    objEngine.StringValue_StringRange(0, "o2") = Loc("DATA_INIZIO") & ": " & StartDate
        
        If SelTable = "WDS_ELAB" Then TipoMedia = MEDIA_ORARIA Else TipoMedia = MEDIA_SEMIORARIA
        
        
        If EndDate >= Now Then EndDate = Now
        
        'Nicolò Maggio 2014: se presenti seleziono la data più vecchia
        'Alby Giugno 2014 datetime invece che date
        'Nicolò Settembre 2013 aggiungo parametro 103 per indicare il formato da gg/mm/aaaa
        strSQL = "SELECT min(" & ParDataQAL2 & ") AS maxDateQAL2 FROM Was_Measures WHERE " & ParDataQAL2 & " <> '' AND C2 IN ('" & Replace(SelParam, ",", "','") & "')"
        .selectionfast (strSQL)
        
        If Not .IsEOF Then
            If .GetValue("maxDateQAL2") <> "" Then
                StartDate = CreateDateFromDB(.GetValue("maxDateQAL2"))
            End If
        End If
        objEngine.StringValue_StringRange(0, "A4") = Replace(Replace(Loc("PERIODO"), "xxxx", Left(StartDate, 10)), "yyyy", Left(EndDate, 10))
        
            RangePar = ""
            For ndx = 0 To MaxPar
                CodPar(ndx, 0) = ListaParametri(ndx)
                
                'Nicola 13/01/2015 copio le celle da C7 a G7 solo dal secondo parametro in poi
                If ndx > 0 Then
                    yDestTopLeft = yDestTopLeft + 5
                    Call objEngine.RangeCopyTo(CoordToString(xStartRange, yStartRange), CoordToString(xEndRange, yEndRange), CoordToString(xStartRange, yDestTopLeft))
                End If
                
                '********** estrae la descrizione par. *********
                strSQL = "SELECT GT_Description FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & CodPar(ndx, 0) & "' AND GT_Description IS NOT NULL"
                .selectionfast (strSQL)
            
                If Not .IsEOF Then CodPar(ndx, 1) = .GetValue("GT_Description")
        
                '********** estrae l'unità di misura ***********
                strSQL = "SELECT GT_Str1 FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & CodPar(ndx, 0) & "' AND GT_Str1 IS NOT NULL"
                .selectionfast (strSQL)
                
                If Not .IsEOF Then CodPar(ndx, 2) = .GetValue("GT_Str1")
                
                'Nicolò Maggio 2014: data inizio QAL2 per singolo parametro.
                
                strSQL = "SELECT " & ParDataQAL2 & " FROM was_measures WHERE C2 = '" & ListaParametri(ndx) & "'"
                .selectionfast (strSQL)
                
                If Not .IsEOF Then
                    'Nicola 18/03/2015 calcolo la data della QAL2, del parametro selezionato con la query precedente, utilizzando la nuova funzione CreateDateFromDB
                    sDataInizioParametro = CreateDateFromDB(.GetValue(ParDataQAL2))
                    CodPar(ndx, 4) = sDataInizioParametro
                End If
                  
                '****** Nome - Unità di misura parametro *****
                objEngine.StringValue_CoordRange(0, 5, 2 + ndx * 5) = CodPar(ndx, 1) & " (" & CodPar(ndx, 2) & ")" & " - (" & CodPar(ndx, 4) & ")"
                
                RangePar = RangePar & "'" & CStr(CodPar(ndx, 0)) & "',"
                
            Next
            
            RangePar = " (" & RangePar & "'" & NFunzPar & "'" & ")"

            'Estrazione del campo di taratura
            .m_Filter = ""
            For i = 0 To MaxPar
                 strSQL = "SELECT " & ParRangeTaratura & " FROM was_measures where C2 = '" & CodPar(i, 0) & "' AND CM_STATIONCODE = '" & SelStation & "'"
                 .selectionfast (strSQL)
                 
                 If (Not .IsEOF) And .GetValue(ParRangeTaratura) <> "" Then
                     RangeTar(i) = .GetValue(ParRangeTaratura)
                 Else
                     RangeTar(i) = -9999
                 End If
            Next

            CheckDate = StartDate
            
            'Nicolò Maggio 2014: seleziono il lunedì successivo, non ha senso considerare dei valori privi della retta o con applicata una retta vecchia
            '**********************************************************************
            '*              posiziona Startdate sul LUNEDI' successivo            *
            '**********************************************************************
            Do While Weekday(StartDate) <> LUNEDI
              StartDate = DateAdd("d", 1, StartDate)
            Loop
            StartWeek = StartDate
            EndWeek = DateAdd("d", 6, StartDate)
            
            '**********************************************************************
            '*            loop fino a che Startdate+6 supera EndDate              *
            '**********************************************************************
            riga = 7  'riga di scrittura dati
            
            yStartRange = 1 'Nicola 14/01/2015 coordinata y della colonna excel della cella da cui parto a copiare la riga
            
            Do While StartWeek < EndDate
            
              '******* reset contatori *******
              MezzoreNormFunz = 0
              For i = 0 To MaxPar
                  ContaOreOKSetteGiorni(i) = 0
                  ContatoreNonOKTarSetteGiorni(i) = 0
              Next
              
              'Nicola - Luca 14/01/2015
              'Implementazione gestione della formattazione delle celle attraverso funzione RangeCopyTo di WindasOcto
              'ATTENZIONE LE FUNZIONI DI WINDASOCTO PARTONO DA 0 COME INDICE -- LA FUNZIONE COORDTOSTRING PARTE DA 1 COME INDICE
              Call objEngine.RangeCopyTo(CoordToString(riga + 1, yStartRange), CoordToString(riga + 1, (yDestTopLeft + 5) - 1), CoordToString(riga + 2, yStartRange))
              
              '********** inserimento periodo **********
               objEngine.StringValue_CoordRange(0, riga, 0) = Left(StartWeek, 10) & "  -  " & Left(EndWeek, 10)
              
              '***** inserimento campo di taratura *****
              For i = 0 To MaxPar
                'Nicolò Maggio 2015: date inserimento rette differenti
                If CDate(CodPar(i, 4)) <= CDate(StartWeek) Then
                    If RangeTar(i) <> -9999 Then
                      objEngine.NumberValue_CoordRange(0, riga, 2 + i * 5) = RangeTar(i)
                    Else
                      objEngine.StringValue_CoordRange(0, riga, 2 + i * 5) = "N.A."
                    End If
                End If
              Next
              
              CheckDate = StartWeek
              
              Dim gg  As Integer
              gg = 0
              
              Do
                  For z = 0 To 47
                      NFunzValue(z) = -9999
                  Next
                   
                  'Estrazione delle medie
                  strSQL = "SELECT * FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' AND dt_measurecod in " & RangePar & _
                           " AND DT_DATE = " & .ParSQLdate(CheckDate) & " ORDER BY DT_HOUR"
                           
                  .m_Filter = ""
                  .selectionfast (strSQL)
                  If (Not .IsEOF) Then
          
                      'controllo se impianto è in normale funzionamento
                      .m_Filter = "DT_MeasureCod = " & .ParSQLStr(NFunzPar)
                      Do While Not rsExtractor.IsEOF
                          z = indiceOra(.GetValue("DT_Hour"), TipoMedia, FORMATO_ORA_IN)
                          NFunzValue(z) = CInt(.GetValue(ParCampoStsImp))
                          .movenext
                      Loop
          
                      'inserimento stato impianto
                      For z = 0 To 47
                          If NFunzValue(z) <> -9999 Then
                              '*** condizione per la definizione di N.Funz.
                              If ImpiantoInMarcia(NFunzValue(z), TipoStatoImpianto, ValoreRifStatoImpianto) Then
                                  MezzoreNormFunz = MezzoreNormFunz + 1
                              End If
                          End If
                      Next
          
                      '***** elaborazione dati per statistiche medie orarie e giornaliere **************
                      For i = 0 To MaxPar
                        'Nicolò Maggio 2014: date inserimento rette differenti
                        If CodPar(i, 4) <= StartWeek Then
                          .m_Filter = "DT_MeasureCod = " & .ParSQLStr(CodPar(i, 0))
                          Do While Not rsExtractor.IsEOF
                          
                              z = indiceOra(.GetValue("DT_Hour"), TipoMedia, FORMATO_ORA_IN)
              
                              If NFunzValue(z) <> -9999 Then
                                  
                                  '************* considera le ore in normale funzionamento **************
                                  If ImpiantoInMarcia(NFunzValue(z), TipoStatoImpianto, ValoreRifStatoImpianto) Then
                                      
                                      '******** considera le ore valide in normale funzionamento ********
                                      valid_flag = "'" & .GetValue(ParCampoValidita) & "'"
                                      If InStr(strValid, valid_flag) > 0 Then
                                          
                                         'contatore di tutte le medie orarie valide per verifica settimanale
                                         ContaOreOKSetteGiorni(i) = ContaOreOKSetteGiorni(i) + 1
                                         
                                         'contatore medie valide NON comprese nel 110% del campo di taratura
                                         If CDbl(RangeTar(i)) > 0 Then
                                            'luca gennaio 2015
                                            misura = .GetValue(ParCampoValore)
                                            If misura > 1.1 * RangeTar(i) Then
                                                ContatoreNonOKTarSetteGiorni(i) = ContatoreNonOKTarSetteGiorni(i) + 1
                                            End If
                                         End If
                                      
                                      End If
                                  End If
                              End If
              
                              .movenext
                          Loop
                        End If
                      Next
          
                   End If
                   
                   gg = gg + 1
                   CheckDate = DateAdd("d", gg, StartWeek)
                  
              Loop While CheckDate <= EndWeek
              
              'calcolo per il 5% e il 40% dei valori > campo di taratura a fine settimana (domenica)
              For i = 0 To MaxPar
                'Nicolò Maggio 2014: date inserimento rette differenti
                If CodPar(i, 4) <= StartWeek Then
                  If ContaOreOKSetteGiorni(i) > 0 Then
                      
                      objEngine.NumberValue_CoordRange(0, riga, 3 + i * 5) = FormatNumber(ContatoreNonOKTarSetteGiorni(i) / ContaOreOKSetteGiorni(i) * 100, 2)
                      objEngine.NumberValue_CoordRange(0, riga, 4 + i * 5) = FormatNumber(ContaOreOKSetteGiorni(i), 0)
                      objEngine.NumberValue_CoordRange(0, riga, 5 + i * 5) = FormatNumber(ContatoreNonOKTarSetteGiorni(i), 0)
                      
                      If ContatoreNonOKTarSetteGiorni(i) / ContaOreOKSetteGiorni(i) * 100 > 5 Then
                          CinquePerCento(i) = CinquePerCento(i) + 1
                      End If
                      
                      If ContatoreNonOKTarSetteGiorni(i) / ContaOreOKSetteGiorni(i) * 100 > 40 Then
                          QuarantaPerCento(i) = QuarantaPerCento(i) + 1
                      End If
                      
                  Else
                      objEngine.StringValue_CoordRange(0, riga, (i + 1) * 3) = "- - -"
                  End If
                End If
              Next
              
              '***** riporto ore di normale funzionamento *****
              objEngine.NumberValue_CoordRange(0, riga, 1) = MezzoreNormFunz
  
              For i = 0 To MaxPar
                'Nicolò Maggio 2014: date inserimento rette differenti
                'Nicola 14/01/2015 aggiungo CDate per vbs altrimenti va in errore
                If CDate(CodPar(i, 4)) <= CDate(StartWeek) Then
                  'indice di disponibilità
                  If MezzoreNormFunz > 0 Then
                      objEngine.StringValue_CoordRange(0, riga, 6 + i * 5) = FormatNumber((ContaOreOKSetteGiorni(i) / MezzoreNormFunz) * 100, 2) & "%"
                  Else
                      objEngine.StringValue_CoordRange(0, riga, 6 + i * 5) = "- - -"
                  End If
                  
                  ContaOreOKSetteGiorni(i) = 0
                  ContatoreNonOKTarSetteGiorni(i) = 0
                End If
              Next
            
              '**** punta alla settimana successiva
              StartWeek = DateAdd("d", 7, StartWeek)
              EndWeek = DateAdd("d", 6, StartWeek)
              
              '**** punta alla riga successiva
              riga = riga + 1
              
            Loop
            
            'luca gennaio 2015 formattazione ultime due righe di riepilogo
            objEngine.RowHeight(riga) = 60
            Call objEngine.RangeCopyTo(CoordToString(riga + 1, 1), CoordToString(riga + 1, yDestTopLeft + 5 - 1), CoordToString(riga + 2, 1))
            Call objEngine.RangeCopyTo(CoordToString(riga + 1, 1), CoordToString(riga + 1, yDestTopLeft + 5 - 1), CoordToString(riga + 3, 1))
            objEngine.RowHeight(riga + 1) = 1000
            objEngine.RowHeight(riga + 2) = 1000
            
            '**** riporta il conteggio delle settimane di supero
            For i = 0 To MaxPar
                'Nicolò Maggio 2014: date inserimento rette differenti
                  If CodPar(i, 4) <= StartWeek Then
                      '**** testo
                      objEngine.StringValue_CoordRange(0, riga + 1, 0) = Loc("STR_QAL2_1")
    
                      '**** numero di settimane con il 5% dei valori > campo di taratura
                      objEngine.NumberValue_CoordRange(0, riga + 1, 3 + i * 5) = CinquePerCento(i)
    
                      '**** testo
                      objEngine.StringValue_CoordRange(0, riga + 2, 0) = Loc("STR_QAL2_2")
    
                      '**** numero di settimane con il 40% dei valori > campo di taratura
                      objEngine.NumberValue_CoordRange(0, riga + 2, 3 + i * 5) = QuarantaPerCento(i)
                  End If
            Next

    End With
    

    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractor = Nothing
    Exit Sub
    
GestErrore:
    Call windasLog("QAL2: " & Error(Err))
    Resume Next
End Sub


Function indiceOra(ByVal Ora, ByVal TipoMedia, ByVal FormatoOraIn)

    On Error GoTo GestErrore

    Const MEDIA_ORARIA = 0
    Const MEDIA_SEMIORARIA = 1

    hh = CInt(Left(Ora, 2))
    mm = CInt(Right(Ora, 2))
    
    If TipoMedia = MEDIA_ORARIA Then
    
      indiceOra = hh
      
    ElseIf TipoMedia = MEDIA_SEMIORARIA Then
    
      If FormatoOraIn = "00.00" Then
        If mm = 0 Then
            indiceOra = hh * 2
        Else
            indiceOra = hh * 2 + 1
        End If
      Else
        indiceOra = (hh * 60 + mm) / 30 - 1 'conta le mezzore
      End If
      
    End If
    
    Exit Function
    
GestErrore:
    Call windasLog("IndiceOra:" & Err())
    
End Function

Private Function ImpiantoInMarcia(ByVal NfunzVal, ByVal TipoStatoImpianto, ByVal ValoreRifStatoImpianto)

    On Error GoTo GestErrore

    Const NF_DA_PERCENTUALE = 0       'assegna Norm. Funz. in base al valore di NFunzPar espresso in percentuale
    Const NF_DA_VALORE = 1            'assegna Norm. Funz. in base al valore di NFunzPar
    
    ImpiantoInMarcia = False
    
    Select Case TipoStatoImpianto
      Case NF_DA_PERCENTUALE
        ImpiantoInMarcia = (NfunzVal >= ValoreRifStatoImpianto)
      Case NF_DA_VALORE
        ImpiantoInMarcia = (NfunzVal = ValoreRifStatoImpianto)
    End Select

    Exit Function
    
GestErrore:
    Call windasLog("ImpiantoInMarcia: " & Err())

End Function

'Nicola - Nicolò 13/01/2015
'Implementazione funzione che date le coordinate x;y di una cella excel ritorna le coordinate alfanumeriche
'Esempio: Passando 1,1 -> A1
Function CoordToString(ByVal riga, ByVal colonna)
       
    On Error GoTo GestErrore
       
    charOfset = 64

    If colonna <= 26 Then
        result = Chr(charOfset + colonna)
    Else
        result = Chr(charOfset + colonna \ 26) & Chr(charOfset + colonna Mod 26)
    End If
    
    result = result & CStr(riga)
    
    CoordToString = result
    
    Exit Function
    
GestErrore:
    Call windasLog("CoordToString: " & Err())
    
End Function



Public Sub QAL3()

    On Error GoTo GestErrore

    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    With rsExtractor
    
    'Nicola 05/03/2015 Scrivo Nome impianto
    objEngine.StringValue_StringRange(0, "F3") = NomeImpianto
    objEngine.StringValue_StringRange(0, "F35") = NomeImpianto
    
    '***** Scrittura della data nell'intestazione *****
    objEngine.StringValue_StringRange(0, "L4") = StartDate
    objEngine.StringValue_StringRange(0, "L36") = StartDate
        
        .selectionfast "SELECT * FROM WDS_gentab where gt_type= 'params' and gt_code ='" & SelParam & "'"
        
        objEngine.StringValue_StringRange(0, "C8") = .GetValue("gt_description")
        objEngine.StringValue_StringRange(0, "C10") = .GetValue("GT_Str1")
        objEngine.StringValue_StringRange(0, "C40") = .GetValue("gt_description")
        objEngine.StringValue_StringRange(0, "C42") = .GetValue("GT_Str1")
        
        
        strSQL = "select * from was_measures where cm_stationcode='" & SelStation & "' and c2='" & SelParam & "'"
        .m_Filter = ""
        .selectionfast (strSQL)
        If Not (.IsEOF) Then

          ZERO_Sams = Replace(.GetValue("c46"), ".", ",")
          SPAN_Sams = Replace(.GetValue("c47"), ".", ",")

          'Costruttore
          objEngine.StringValue_StringRange(0, "C5") = .GetValue("C48")
          objEngine.StringValue_StringRange(0, "C37") = .GetValue("C48")
          'Analizzatore
          objEngine.StringValue_StringRange(0, "C6") = .GetValue("C49")
          objEngine.StringValue_StringRange(0, "C38") = .GetValue("C49")
          'Serial number
          objEngine.StringValue_StringRange(0, "C7") = .GetValue("C50")
          objEngine.StringValue_StringRange(0, "C39") = .GetValue("C50")
          'Campo Misura
          objEngine.StringValue_StringRange(0, "C9") = "0-" & .GetValue("C9")
          objEngine.StringValue_StringRange(0, "C41") = "0-" & .GetValue("C9")
          'Numero certificato
          objEngine.StringValue_StringRange(0, "C11") = .GetValue("C52")
          objEngine.StringValue_StringRange(0, "C43") = .GetValue("C52")
        End If
        
        '****************** Estrazione dati da tabella ******************
        strSQL = "SELECT * FROM WDS_CALIBRATION"
        strSQL = strSQL & " WHERE CL_STATION = '" & SelStation & "' AND CL_PARAMETER = '" & SelParam & "'"
        strSQL = strSQL & " AND CL_DATE<=" & .ParSQLdate(StartDate) & ""
        strSQL = strSQL & " ORDER BY cl_date DESC, cl_hour DESC"

        .m_Filter = ""
        .selectionfast (strSQL)
        If Not (.IsEOF) Then
        
            '*************** nel foglio VERIFICA DI PRECISIONE **************
            If .RecordCount = 1 Then
              
              objEngine.StringValue_StringRange(0, "J18") = "0"
              objEngine.StringValue_StringRange(0, "L18") = "0"
            End If

            .MoveFirst
                       
                    
            '****************************************************************
            '*                 TABELLA VERIFICA PRECISIONE                  *
            '****************************************************************
                      
                    '****** ZERO - verifica attuale *****
            objEngine.StringValue_StringRange(0, "C19") = Mid(.GetValue("cl_date"), 7, 2) & "/" & Mid(.GetValue("cl_date"), 5, 2) & "/" & Mid(.GetValue("cl_date"), 1, 4) & " " & .GetValue("cl_hour")

            objEngine.NumberValue_StringRange(0, "D19") = Replace(.GetValue("cl_tspan4"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "E19") = Replace(.GetValue("cl_tspan2"), ".", ",")
                                
            objEngine.NumberValue_StringRange(0, "F19") = Replace(.GetValue("c5"), ".", ",")
                                
            objEngine.NumberValue_StringRange(0, "G19") = ZERO_Sams

            objEngine.NumberValue_StringRange(0, "H19") = Replace(.GetValue("c1"), ".", ",")

            objEngine.NumberValue_StringRange(0, "I19") = Replace(.GetValue("c3"), ".", ",")
    
            objEngine.NumberValue_StringRange(0, "J19") = Replace(.GetValue("c9"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "K19") = Replace(.GetValue("c7"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "L19") = Replace(.GetValue("c11"), ".", ",")
            
            objEngine.StringValue_StringRange(0, "M19") = Loc(.GetValue("c13"))
            
            '****** SPAN - verifica attuale *****
            objEngine.StringValue_StringRange(0, "C23") = Mid(.GetValue("cl_date"), 7, 2) & "/" & Mid(.GetValue("cl_date"), 5, 2) & "/" & Mid(.GetValue("cl_date"), 1, 4) & " " & .GetValue("cl_hour")
            
            objEngine.NumberValue_StringRange(0, "D23") = Replace(.GetValue("cl_tspan5"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "E23") = Replace(.GetValue("cl_tspan3"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "F23") = Replace(.GetValue("c6"), ".", ",")
                                
            objEngine.NumberValue_StringRange(0, "G23") = SPAN_Sams
            
            objEngine.NumberValue_StringRange(0, "H23") = Replace(.GetValue("c2"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "I23") = Replace(.GetValue("c4"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "J23") = Replace(.GetValue("c10"), ".", ",")
                                
            objEngine.NumberValue_StringRange(0, "K23") = Replace(.GetValue("c8"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "L23") = Replace(.GetValue("c12"), ".", ",")
            
            objEngine.StringValue_StringRange(0, "M23") = Loc(.GetValue("c14"))
                    
                    '********* punta alla verifica precedente *******
            If .RecordCount > 1 Then
                .movenext
                
                '****** ZERO - verifica precedente *****
                objEngine.StringValue_StringRange(0, "C18") = Mid(.GetValue("cl_date"), 7, 2) & "/" & Mid(.GetValue("cl_date"), 5, 2) & "/" & Mid(.GetValue("cl_date"), 1, 4) & " " & .GetValue("cl_hour")
                
                objEngine.NumberValue_StringRange(0, "D18") = Replace(.GetValue("cl_tspan4"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "E18") = Replace(.GetValue("cl_tspan2"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "F18") = Replace(.GetValue("c5"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "G18") = ZERO_Sams
                
                objEngine.NumberValue_StringRange(0, "H18") = Replace(.GetValue("c1"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "I18") = Replace(.GetValue("c3"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "J18") = Replace(.GetValue("c9"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "K18") = Replace(.GetValue("c7"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "L18") = Replace(.GetValue("c11"), ".", ",")
                
                objEngine.StringValue_StringRange(0, "M18") = Loc(.GetValue("c13"))
                
                '****** SPAN - verifica precedente *****
                objEngine.StringValue_StringRange(0, "C22") = Mid(.GetValue("cl_date"), 7, 2) & "/" & Mid(.GetValue("cl_date"), 5, 2) & "/" & Mid(.GetValue("cl_date"), 1, 4) & " " & .GetValue("cl_hour")
                
                objEngine.NumberValue_StringRange(0, "D22") = Replace(.GetValue("cl_tspan5"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "E22") = Replace(.GetValue("cl_tspan3"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "F22") = Replace(.GetValue("c6"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "G22") = SPAN_Sams
                
                objEngine.NumberValue_StringRange(0, "H22") = Replace(.GetValue("c2"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "I22") = Replace(.GetValue("c4"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "J22") = Replace(.GetValue("c10"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "K22") = Replace(.GetValue("c8"), ".", ",")
                
                objEngine.NumberValue_StringRange(0, "L22") = Replace(.GetValue("c12"), ".", ",")
                
                objEngine.StringValue_StringRange(0, "M22") = Loc(.GetValue("c14"))

            End If
                    
            '****************************************************************
            '*                   TABELLA VERIFICA DERIVA                    *
            '****************************************************************
            
            .MoveFirst
            
            '****** TABELLA DI ZERO ******
            objEngine.NumberValue_StringRange(0, "B49") = ZERO_Sams
            
            objEngine.NumberValue_StringRange(0, "D49") = Replace(.GetValue("c20"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "F49") = Replace(.GetValue("c22"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "C50") = Replace(.GetValue("cl_tspan2"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "B52") = Replace(.GetValue("cl_tspan4"), ".", ",")
            
            If .RecordCount = 1 Then
              objEngine.StringValue_StringRange(0, "D52") = "0"
              
              objEngine.StringValue_StringRange(0, "D53") = "0"
              
              objEngine.StringValue_StringRange(0, "F52") = "0"
              
              objEngine.StringValue_StringRange(0, "F53") = "0"
            End If
            
            objEngine.NumberValue_StringRange(0, "C54") = Replace(.GetValue("c24"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "B56") = Replace(.GetValue("c26"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "E56") = Replace(.GetValue("c29"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "B59") = Replace(.GetValue("c27"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "B60") = Replace(.GetValue("c30"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "E59") = Replace(.GetValue("c28"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "E60") = Replace(.GetValue("c31"), ".", ",")
            
            '*** esito
            objEngine.StringValue_StringRange(0, "A62") = Loc(.GetValue("c38"))
            
            If objEngine.StringValue_StringRange(0, "A62") = Loc("DERIVA POSITIVA") Then
              objEngine.NumberValue_StringRange(0, "C63") = Replace(.GetValue("c40"), ".", ",")
            ElseIf objEngine.StringValue_StringRange(0, "A62") = Loc("DERIVA NEGATIVA") Then
              objEngine.NumberValue_StringRange(0, "C63") = Replace(.GetValue("c41"), ".", ",")
            End If
            
            
            '****** TABELLA DI SPAN ******
            objEngine.NumberValue_StringRange(0, "J49") = SPAN_Sams
            
            objEngine.NumberValue_StringRange(0, "L49") = Replace(.GetValue("c21"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "N49") = Replace(.GetValue("c23"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "K50") = Replace(.GetValue("cl_tspan3"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "J52") = Replace(.GetValue("cl_tspan5"), ".", ",")
            
            If .RecordCount = 1 Then
              objEngine.StringValue_StringRange(0, "L52") = "0"
              
              objEngine.StringValue_StringRange(0, "L53") = "0"
              
              objEngine.StringValue_StringRange(0, "N52") = "0"
              
              objEngine.StringValue_StringRange(0, "N53") = "0"
            End If
            
            objEngine.NumberValue_StringRange(0, "K54") = Replace(.GetValue("c25"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "J56") = Replace(.GetValue("c32"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "M56") = Replace(.GetValue("c35"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "J59") = Replace(.GetValue("c33"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "J60") = Replace(.GetValue("c36"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "M59") = Replace(.GetValue("c34"), ".", ",")
            
            objEngine.NumberValue_StringRange(0, "M60") = Replace(.GetValue("c37"), ".", ",")
            
            '*** esito
            objEngine.StringValue_StringRange(0, "I62") = Loc(.GetValue("c39"))
            
            If objEngine.StringValue_StringRange(0, "I62") = Loc("DERIVA POSITIVA") Then
              objEngine.NumberValue_StringRange(0, "K63") = Replace(.GetValue("c42"), ".", ",")
            ElseIf objEngine.StringValue_StringRange(0, "I62") = Loc("DERIVA NEGATIVA") Then
              objEngine.NumberValue_StringRange(0, "K63") = Replace(.GetValue("c43"), ".", ",")
            End If
            
            '********* punta alla verifica precedente *******
            If .RecordCount > 1 Then
              .movenext

              '****** TABELLA DI ZERO ******
              objEngine.NumberValue_StringRange(0, "D52") = Replace(.GetValue("c27"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "D53") = Replace(.GetValue("c30"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "F52") = Replace(.GetValue("c28"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "F53") = Replace(.GetValue("c31"), ".", ",")
              
              '****** TABELLA DI SPAN ******
              objEngine.NumberValue_StringRange(0, "L52") = Replace(.GetValue("c33"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "L53") = Replace(.GetValue("c36"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "N52") = Replace(.GetValue("c34"), ".", ",")
              
              objEngine.NumberValue_StringRange(0, "N53") = Replace(.GetValue("c37"), ".", ",")
              
            End If
        
        End If

    End With
    
    '*********** generazione tabella finale **************
    objEngine.Save2Xls ("c:\windas\temp.xls")
    Set rsExtractor = Nothing
    Set objEngine = Nothing
    
    Exit Sub
    
GestErrore:
    Call windasLog("QAL3: " & Error(Err))
    
End Sub



Public Property Let DbType(x As String)
    m_DbType = x
End Property
Public Property Get DbType() As String
    DbType = m_DbType
End Property

Public Property Let DbDatabase(x As String)
    m_DbDatabase = x
End Property
Public Property Get DbDatabase() As String
    DbDatabase = m_DbDatabase
End Property

Public Property Let DbUser(x As String)
    m_DbUser = x
End Property
Public Property Get DbUser() As String
    DbUser = m_DbUser
End Property

Public Property Let DbPassword(x As String)
    m_DbPassword = x
End Property
Public Property Get DbPassword() As String
    DbPassword = m_DbPassword
End Property

Public Property Let DbServer(x As String)
    m_DbServer = x
End Property
Public Property Get DbServer() As String
    DbServer = m_DbServer
End Property

Public Property Let DbVersion(x As String)
    m_DbVersion = x
End Property
Public Property Get DbVersion() As String
    DbVersion = m_DbVersion
End Property

Public Sub SetStartDateFromString(StringDate As String)
    'Considerando che la data passata da BFDesk o WinCC_PrENG o BFNotify è nel formato gg/mm/aaaa etc , creo la data in modo indipendente dalle impostazione internazionali
    'Nicolò Agosto 2015 Gestisco anche ore
    If Len(StringDate) > 10 Then
        StartDate = DateTimeSerial(CInt(Mid(StringDate, 7, 4)), CInt(Mid(StringDate, 4, 2)), CInt(Mid(StringDate, 1, 2)), CInt(Mid(StringDate, 12, 2)), CInt(Mid(StringDate, 15, 2)), 0)
    Else
        StartDate = DateSerial(CInt(Mid(StringDate, 7, 4)), CInt(Mid(StringDate, 4, 2)), CInt(Mid(StringDate, 1, 2)))
    End If
    'OLD: StartDate = DateSerial(CInt(Mid(StringDate, 7, 4)), CInt(Mid(StringDate, 4, 2)), CInt(Mid(StringDate, 1, 2)))
End Sub
Public Property Let StartDate(x As Date)
    m_StartDate = x
End Property
Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Sub SetEndDateFromString(StringDate As String)
    'Considerando che la data passata da BFDesk o WinCC_PrENG o BFNotify è nel formato gg/mm/aaaa etc , creo la data in modo indipendente dalle impostazione internazionali
    EndDate = DateSerial(CInt(Mid(StringDate, 7, 4)), CInt(Mid(StringDate, 4, 2)), CInt(Mid(StringDate, 1, 2)))
End Sub
Public Property Let EndDate(x As Date)
    m_EndDate = x
End Property
Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Let Stations(x As String)
    m_Stations = x
End Property
Public Property Get Stations() As String
    Stations = m_Stations
End Property

Public Property Let Param(x As String)
    m_Param = x
End Property
Public Property Get Param() As String
    Param = m_Param
End Property

Public Property Let Table(x As String)
    m_Table = x
End Property
Public Property Get Table() As String
    Table = m_Table
End Property

Public Property Let ModelFileName(x As String)
    m_ModelFileName = x
End Property
Public Property Get ModelFileName() As String
    ModelFileName = m_ModelFileName
End Property


Public Sub DefaultConnections()
    DbDatabase = "bfdata"
    DbUser = "root"
    DbPassword = "Bfinfo9876"
    DbServer = "127.0.0.1"
    DbVersion = "MYSQL5"
    StartDate = "04/03/2015 18:00"
    EndDate = "04/03/2015 18:00"
    Stations = "CEMS"
    Param = "CO,H2O"
    Table = "wds_elab"
End Sub

Private Function GetStsImpDescription(codstsimp As Integer) As String
    'sono accettati gli stati impianto nell'intervallo ]0-100[ estremi esclusi
    On Error GoTo GestErrore
    
    Dim result As String
    result = Loc("STR_IMP_UNDEFINED")
    
    If codstsimp > 0 And codstsimp < 100 Then
         result = Loc("STR_IMP_" & Format(codstsimp, "00"))
    End If
    
    GetStsImpDescription = result
    
    Exit Function
GestErrore:
   Call windasLog("GetStsImpDescription" & Error(Err))
   GetStsImpDescription = result
End Function
Private Function RicavaCodiceImpianto(PrefissoStazione As Integer, IsGruppi As Boolean) As Integer

On Error GoTo GestErrore
    
    If Not IsGruppi Then
        Select Case PrefissoStazione
            Case 1
                RicavaCodiceImpianto = 1
            Case 2
                RicavaCodiceImpianto = 2
            Case 3
                RicavaCodiceImpianto = 3
            Case 4
                RicavaCodiceImpianto = 5
            Case Else
                RicavaCodiceImpianto = 1
        End Select
    Else
        Select Case PrefissoStazione
            Case 1
                RicavaCodiceImpianto = 4
            Case 4
                RicavaCodiceImpianto = 6
            Case Else
                RicavaCodiceImpianto = 4
        End Select
    End If
        
Exit Function

GestErrore:
   Call windasLog("RicavaCodiceImpianto" & Error(Err))
End Function
Sub SostituisciFlagPiemonte(StrRiga)

    On Error GoTo GestErrore
    
    If IsPiemonte Then
        If InStr(UCase(StrRiga), "VAL") > 0 Then StrRiga = Replace(StrRiga, "VAL", "V")
        If InStr(UCase(StrRiga), "ERR") > 0 Then StrRiga = Replace(StrRiga, "ERR", "AL")
        If InStr(UCase(StrRiga), "TSP") > 0 Then StrRiga = Replace(StrRiga, "TSP", "TAR")
        If InStr(UCase(StrRiga), "TZR") > 0 Then StrRiga = Replace(StrRiga, "TZR", "TAR")
        If InStr(UCase(StrRiga), "NVA") > 0 Then StrRiga = Replace(StrRiga, "NVA", "CRT")
        If InStr(UCase(StrRiga), "NVH") > 0 Then StrRiga = Replace(StrRiga, "NVH", "OVRANGE")
        If InStr(UCase(StrRiga), "NVL") > 0 Then StrRiga = Replace(StrRiga, "NVL", "UNRANGE")
        If InStr(UCase(StrRiga), "NCO") > 0 Then StrRiga = Replace(StrRiga, "NCO", "ALTRO")
        If InStr(UCase(StrRiga), "NCU") > 0 Then StrRiga = Replace(StrRiga, "NCU", "ALTRO")
        If InStr(UCase(StrRiga), "OFF") > 0 Then StrRiga = Replace(StrRiga, "OFF", "ALTRO")
    Else
        StrRiga = ""
    End If

    Exit Sub
    
GestErrore:
    Call windasLog("SostituisciFlagPiemonte: " & Error(Err))

End Sub

'luca novembre 2018
Public Sub Totalizzazioni(strIndiciDigitali As String)

    On Error GoTo GestErrore

    Dim TotaleBypassOre As Integer
    Dim TotaleBypassSecondi As Integer
    Dim anno As Integer
  
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    
    If objEngine Is Nothing Then
        Call windasLog("Totalizzazioni: Impossibile creare l'oggetto: WindasOcto.CWindas_Spread")
        Exit Sub
    End If
    
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.Worksheet = 0
    
    With rsExtractor
        
        anno = year(StartDate)
        objEngine.StringValue_StringRange(0, "B6") = CStr(anno)
        objEngine.StringValue_StringRange(0, "B2") = NomeImpianto
        objEngine.StringValue_StringRange(1, "C2") = NomeImpianto
        objEngine.StringValue_StringRange(1, "G2") = "Anno: " & CStr(anno)
        
        '*******************************************************************************************
        '*              estrazione totale ore impianto in bypass                   *
        '*******************************************************************************************
        Dim DataIn As String
        Dim dDataIn As Date
        Dim OraIn As String
        Dim DataFin As String
        Dim dDataFin As Date
        Dim OraFin As String
        Dim TotSecondi As Long
        Dim secondi As Long
        
        '***** riga iniziale
        RigaAlarm = 0
        
        '***** estrazione eventi relativi al bypass
        strSQL = "SELECT * FROM WDS_ALARM WHERE AL_STATION = '" & SelStation & "'"
        strSQL = strSQL & " AND Al_Parameter in ('" & strIndiciDigitali & "')"
        strSQL = strSQL & " AND left(AL_DATE, 4) = '" & anno & "'"
        strSQL = strSQL & " order by AL_date, AL_hour"

        .m_Filter = ""
        .selectionfast (strSQL)
        Do While Not .IsEOF
            '***** Riporta evento *****
            objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 2) = .GetValue("AL_DESCRIPTION")

            '***** Inizio *****
            DataIn = .GetValue("AL_date")
            DataIn = Mid(DataIn, 7, 2) & "/" & Mid(DataIn, 5, 2) & "/" & Mid(DataIn, 1, 4)
            OraIn = .GetValue("AL_HOUR")
            objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 3) = DataIn
            objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 4) = OraIn

            '***** Fine *****
            DataFin = .GetValue("AL_date2")
            If DataFin <> "" Then DataFin = Mid(DataFin, 7, 2) & "/" & Mid(DataFin, 5, 2) & "/" & Mid(DataFin, 1, 4)
            OraFin = .GetValue("AL_HOUR2")
            objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 5) = DataFin
            objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 6) = OraFin
            .movenext

            '***** calcola durata evento
            If IsDate(DataIn & " " & OraIn) And IsDate(DataFin & " " & OraFin) Then
                dDataIn = CDate(DataIn & " " & OraIn)
                dDataFin = CDate(DataFin & " " & OraFin)
                If dDataFin >= dDataIn Then
                    secondi = DateDiff("s", dDataIn, dDataFin)
                    TotSecondi = TotSecondi + secondi
                    objEngine.NumberValue_CoordRange(1, RigaAlarm + 6, 7) = secondi
                End If
            End If

            RigaAlarm = RigaAlarm + 1
        Loop
        TotaleBypassOre = TotSecondi \ 3600
        TotaleBypassSecondi = TotSecondi Mod 3600
        
    End With

    objEngine.StringValue_StringRange(0, "C6") = TotaleBypassOre
    objEngine.StringValue_StringRange(0, "D6") = TotaleBypassSecondi

    objEngine.Save2Xls ("C:\Windas\Temp.xls")
    
    Set rsExtractor = Nothing
    Set rsNfunz = Nothing

Exit Sub
    
GestErrore:
    Call windasLog("Totalizzazioni: " & Error(Err))
    Resume Next

End Sub

