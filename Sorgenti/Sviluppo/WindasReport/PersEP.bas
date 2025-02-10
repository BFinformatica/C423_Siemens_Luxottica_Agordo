Attribute VB_Name = "PersEP"
'Federica giugno 2017 report mensile personalizzato EP
Public Sub MensileEP(ByRef NFunzValue() As Double, ByRef MinVal() As Double, ByRef MaxVal() As Double)

    Dim Value As String
    Dim Validflag As String
    Dim StatoImpianto As String 'Non serve più
    Dim nn As Integer
    Dim campoNumCampioni As String
    Dim Parametri(8, 2) As String
    Dim ValueGrezzo As String
    Dim ValidFlagGrezzo As String
    Dim campoNumCampioniGrezzo As String
    Dim ValueFlusso As String
    Dim GiorniInvalidi(8, 31) As Boolean
    Dim NumeroSuperamenti(8) As Integer
    Dim OreEstratte As Integer
    'luca 22/11/2016
    Dim Ore_FS As Integer
    Dim Ore_AS_CA As Integer
    
    Dim rsExtractorLocale As Object
    
    On Error GoTo GestErrore
    
    Value = "dt_value"
    Validflag = "dt_validflag"
    campoNumCampioni = "DT_NR" 'Nicolò Settembre 2015
    ValueGrezzo = "dt_valuetq"
    ValidFlagGrezzo = "dt_validflag_tq"
    campoNumCampioniGrezzo = "dt_nr_tq"
    ValueFlusso = "dt_fm"
    StatoImpianto = "DT_CUSTOM1"
    
    Set rsExtractorLocale = CreateObject("AttimoFwk.CData")
    With rsExtractorLocale
        Call CReport.InizializzaExtractor(rsExtractorLocale)
'        Select Case CReport.DbType
'            Case "MYSQL5"
'                .SetDBType .Conn_MYSQL5
'                rsNfunz.SetDBType .Conn_MYSQL5
'            Case "MYSQL"
'                .SetDBType .Conn_MYSQL
'                rsNfunz.SetDBType .Conn_MYSQL
'            Case "ORACLE"
'                .SetDBType .Conn_Oracle
'                rsNfunz.SetDBType .Conn_Oracle
'            Case "SQL"
'                .SetDBType .Conn_SQL
'                rsNfunz.SetDBType .Conn_SQL
'        End Select
'
'        '***** Database, utente e password *****
'        .SetDatabase DbDatabase, DbUser, DbPassword
'        rsNfunz.SetDatabase DbDatabase, DbUser, DbPassword
'        '***** Server ******
'        .SetServer DbServer
'        rsNfunz.SetServer DbServer
'        '***** DB version ******
'        .SetDbVersion DbVersion
'        rsNfunz.SetDbVersion DbVersion
    End With
    
    Set objEngine = CreateObject("WindasOcto.CWindas_Spread")
    objEngine.OpenFile ("" & ModelFileName)
    objEngine.WorkSheet = 0
    
    'Report con colonne contenenti diverse informazioni - cablo i parametri
    Parametri(0, 0) = "CME" 'CARICO MASSIMO ESPRIMIBILE
    Parametri(1, 0) = "MTA" 'MINIMO TECNICO AMBIENTALE
    Parametri(2, 0) = "CET" 'CARICO EFFETTIVO TURBINA
    Parametri(3, 0) = "QGas" 'PORTATA COMBUSTIBILE
    Parametri(4, 0) = "Qfumi"   'PORTATA FUMI
    Parametri(5, 0) = "Tfumi"   'TEMPERATURA
    Parametri(6, 0) = "O2"  'O2
    Parametri(7, 0) = "CO"  'CO
    Parametri(8, 0) = "NOx" 'NOx
    
    MaxPar = UBound(Parametri, 1)
    
    With rsExtractor
        
        '***** Scrittura nome impianto *****
        objEngine.StringValue_StringRange(0, "B2") = NomeImpianto
'        objEngine.StringValue_StringRange(1, "C2") = NomeImpianto
'
        '***** Scrittura della data nell'intestazione *****
        objEngine.StringValue_StringRange(0, "AE2") = Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy")
'        objEngine.StringValue_StringRange(1, "F2") = Loc("DATA_INIZIO") & ": " & Format(StartDate, "dd/mm/yyyy")
'
        'scrivo O2 Riferimento
        If Val(O2_rif) > 0 Then
            objEngine.StringValue_StringRange(0, "Q52") = "Le misure di emissione sono riferite ad un tenore di ossigeno del " & O2_rif & "% Vol."
        End If

        '***** Azzeramento matrici *****
        For z = 0 To 31
            NFunzValue(z) = -9999
        Next

        For i = 0 To MaxPar
            MinVal(i, nn) = 999999
            MaxVal(i, nn) = -999999
        Next

        OreEstratte = 0
        'Ore per stato impianto
        Ore_AS_CA = 0
        Ore_FS = 0

        '***** Ciclo per parametro *****
        For i = 0 To MaxPar

            'estraggo descrizione del parametro
            strSQL = "SELECT GT_Description FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & Parametri(i, 0) & "' AND GT_Description IS NOT NULL"
            .selectionfast strSQL
            If Not .IsEOF Then Parametri(i, 1) = .getValue("GT_Description")

            'estraggo unità di misura del parametro
            strSQL = "SELECT GT_Str1 FROM wds_GenTab WHERE GT_Type = 'PARAMS' AND GT_Code = '" & Parametri(i, 0) & "' AND GT_Str1 IS NOT NULL"
            .selectionfast strSQL
            If Not .IsEOF Then Parametri(i, 2) = .getValue("GT_Str1")

            '***** Scrittura del Nome parametro e unità di misura nell'intestazione *****
            objEngine.StringValue_CoordRange(0, 5, CReport.RicavaPosizioneExcel(i)) = Parametri(i, 1) & " (" & Parametri(i, 2) & ")"
        Next i

        '***** Estrazione medie giornaliere *****
        strSQL = "SELECT * FROM WDS_DAYS WHERE " & _
                " DT_STATIONCODE = '" & SelStation & "' AND DT_DATE BETWEEN '" & String_StartDate_DB & "' AND '" & String_EndDate_DB & "' ORDER BY DT_DATE"
        .selectionfast strSQL

        If Not .IsEOF Then

            '***** Filtro per stato impianto *****
            .m_filter = "DT_MEASURECOD = '" & NFunzPar & "'"
            '***** Caricamento in matrice dei valori di stato impianto *****
            Do While Not .IsEOF
                NFunzValue(Right(.getValue("DT_DATE"), 2)) = .getValue("DT_NR") 'ho già le ore di funzionamento normale
                OreEstratte = OreEstratte + .getValue("DT_NR")
                .movenext
            Loop
            .m_filter = ""

            '***** Inserimento stato impianto orario *****
            'Ore di normale funzionamento nel mese
            strSQL = "SELECT SUM(DT_NR) AS TotOre FROM WDS_DAYS " & _
                     "WHERE DT_STATIONCODE = '" & SelStation & "' and DT_DATE between '" & String_StartDate_DB & "' and '" & String_EndDate_DB & "' " & _
                     "AND DT_MEASURECOD = '" & NFunzPar & "'"
            rsExtractorLocale.selectionfast strSQL
            OreNormFunz = rsExtractorLocale.getValue("TotOre")
            objEngine.StringValue_StringRange(0, "H53") = OreNormFunz
            
            'Ore CA o AS nel mese
            strSQL = "SELECT COUNT(*) AS TotOre FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' " & _
                     "AND DT_DATE between '" & String_StartDate_DB & "' and '" & String_EndDate_DB & "' " & _
                     "AND DT_MEASURECOD = '" & NFunzPar & "' and DT_CUSTOM1 IN('31', '36')"
            rsExtractorLocale.selectionfast strSQL
            Ore_AS_CA = rsExtractorLocale.getValue("TotOre")
            objEngine.StringValue_StringRange(0, "H52") = Ore_AS_CA
            
            'Ore FS nel mese
            strSQL = "SELECT COUNT(*) AS TotOre FROM " & SelTable & " WHERE DT_STATIONCODE = '" & SelStation & "' " & _
                     "AND DT_DATE between '" & String_StartDate_DB & "' and '" & String_EndDate_DB & "' " & _
                     "AND DT_MEASURECOD = '" & NFunzPar & "' and DT_CUSTOM1 IN('34')"
            rsExtractorLocale.selectionfast strSQL
            Ore_FS = rsExtractorLocale.getValue("TotOre")
            objEngine.StringValue_StringRange(0, "H51") = Ore_FS
            
            For z = 1 To 31
                If NFunzValue(z) <> -9999 Then
                    objEngine.StringValue_CoordRange(0, z + 7, 3) = NFunzValue(Format(z, "00"))
                End If
            Next z

            '***** Ciclo per parametro *****
            For i = 0 To MaxPar
                Call CReport.MensileEPNormalizzati(rsExtractor, Parametri(i, 0), Validflag, Value, campoNumCampioni, GiorniInvalidi, NumeroSuperamenti)
                'solo per CO e NOx visualizzo medie giornaliere grezze
                If i >= 7 Then
                    strSQL = "SELECT dt_date, AVG(DT_VALUETQ) AS MediaGiorno FROM WDS_elab " & _
                             "WHERE DT_STATIONCODE = '" & SelStation & "' and DT_DATE between '" & String_StartDate_DB & "' and '" & String_EndDate_DB & "' " & _
                             "AND DT_MEASURECOD = '" & Parametri(i, 0) & "' " & _
                             "GROUP BY dt_date ORDER BY DT_DATE"
                    rsExtractorLocale.selectionfast strSQL
                    Do While Not rsExtractorLocale.IsEOF
                        CurRow = (Right(rsExtractorLocale.getValue("DT_DATE"), 2)) + 7
                        objEngine.NumberValue_CoordRange(0, CurRow, CReport.RicavaPosizioneExcel(i)) = rsExtractorLocale.getValue("MediaGiorno")
                        
                        rsExtractorLocale.movenext
                    Loop
                    
                    'Media mensile sui dati normnalizzati
                    strSQL = "SELECT DT_VALUE FROM wds_month " & _
                             "WHERE DT_STATIONCODE = '" & SelStation & "' and DT_DATE = '" & Left(String_StartDate_DB, 6) & "' " & _
                             "AND DT_MEASURECOD = '" & Parametri(i, 0) & "' "
                    rsExtractorLocale.selectionfast strSQL
                    If Not rsExtractorLocale.IsEOF Then
                        objEngine.NumberValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 2) = rsExtractorLocale.getValue("DT_VALUE")
                    End If
                End If

            Next
            .m_filter = ""
        End If

        '***** Ciclo per parametro *****
        For i = 7 To MaxPar
            If Len(Parametri(i, 0)) > 0 Then
                '***** Se ci sono almeno 6 ore di normal funzionamento *****
                If CDbl(OreNormFunz) > 144 Then
                    '***** Se c'è almeno una media giornaliera valida *****
                    If CReport.ContaValidi(i, nn) > 0 Then
                        '***** Status di validità media giornaliera *****
                        
                        strSQL = "SELECT (DT_NR / DT_NRTOT) AS Indice, DT_VALIDFLAG FROM WDS_MONTH " & _
                                 "WHERE DT_DATE = '" & Left(String_StartDate_DB, 6) & "' AND DT_MEASURECOD = '" & Parametri(i, 0) & "' "
                        rsExtractorLocale.selectionfast strSQL
                        If Not rsExtractor.IsEOF Then
                        
                            '***** Media giornaliera valida solo con ID% >= 70% *****
                            If (rsExtractorLocale.getValue("DT_VALIDFLAG") <> "VAL") And (rsExtractorLocale.getValue("DT_VALIDFLAG") <> "AUX") Then
                                objEngine.StringValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 1) = "*"
                            End If
                            
                            '***** ID% media mensile *****
                            objEngine.NumberValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 3) = rsExtractorLocale.getValue("Indice")
                       
                        End If

                    Else
                        '***** Status di validità media giornaliera - non calcolata *****
                        'luca 11/11/2016 su richiesta del cliente media giorno solo per CO e NOx
                        If i >= 7 Then
                            objEngine.StringValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 1) = "*"
                        End If

                        '***** ID% media giornaliera *****
                        'luca 11/11/2016 su richiesta del cliente media giorno solo per CO e NOx
                        If i >= 7 Then
                            objEngine.NumberValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 3) = 0
                        End If
                    End If
                Else
                    '***** Status di validità media giornaliera - non significativa*****
                    'luca 11/11/2016 su richiesta del cliente media giorno solo per CO e NOx
                    If i >= 7 Then
                        objEngine.StringValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 1) = "*"
                    End If
                    '***** ID% media giornaliera *****
                    'luca 11/11/2016 su richiesta del cliente media giorno solo per CO e NOx
                    If i >= 7 Then
                        objEngine.NumberValue_CoordRange(0, 44, CReport.RicavaPosizioneExcel(i) + 3) = 0
                    End If
                End If
            End If
        Next
'
'        '****** Inserimento numero ore di normale funzionamento giornaliere ******
'        'luca 11/11/2016 disabilito
'        'objEngine.NumberValue_CoordRange(0, 37, 34) = OreNormFunz
'
        Call CReport.MensileEPOreValidato(rsExtractorLocale, Parametri())
'
'        '***** Inserimento allarmi con riconoscimento *****
'        strSQL = "SELECT * FROM WDS_ALARM WHERE AL_STATION = '" & SelStation & "' AND AL_DATE = '" & String_StartDate_DB & "' ORDER BY AL_DATE, AL_HOUR, AL_DESCRIPTION" 'Nicolò ordino anche per data e ora
'
'        If (.selectionfast(strSQL)) Then
'            Do While Not rsExtractor.isEOF
'
'                 '***** Allarme *****
'                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 2) = .getValue("AL_DESCRIPTION") & " " & .getValue("AL_STATUSDESC")
'
'                 '*************************** inizio allarme **************************
'                 '***** Data *****
'                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 3) = Right(.getValue("AL_DATE"), 2) & "/" & Mid(.getValue("AL_DATE"), 5, 2) & "/" & Left(.getValue("AL_DATE"), 4)
'                 '***** Ora *****
'                 objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 4) = .getValue("AL_HOUR")
'
'                 '*************************** rientro allarme **************************
'                 '***** Data *****
'                 If .getValue("AL_DATE2") <> "" Then
'                    objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 5) = Right(.getValue("AL_DATE2"), 2) & "/" & Mid(.getValue("AL_DATE2"), 5, 2) & "/" & Left(.getValue("AL_DATE2"), 4)
'                    '***** Ora *****
'                    objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 6) = .getValue("AL_HOUR2")
'                 Else
'                    objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 5) = "---"
'                    '***** Ora *****
'                    objEngine.StringValue_CoordRange(1, RigaAlarm + 6, 6) = "---"
'
'                 End If
'                 RigaAlarm = RigaAlarm + 1
'
'                 .movenext
'            Loop
'        End If

    End With

    objEngine.Save2Xls ("C:\Windas\Temp.xls")
 
    Set objEngine = Nothing
    Set rsExtractorLocale = Nothing
    Set rsExtractor = Nothing
    
    Exit Sub
    
GestErrore:
    Call CReport.windasLog("MensileEP: " & Error(Err))
    Resume Next
End Sub


