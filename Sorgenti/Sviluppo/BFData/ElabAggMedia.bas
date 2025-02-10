Attribute VB_Name = "ElabAggMedia"
Option Explicit

Sub ElaboraAggiornaMedia(periodo As Integer, tipo As Integer, iIdx1 As Integer, Elabdate, tipodato As Integer, numMedia As Integer)

    Dim Windas As Object
    Dim rsDati As Object
    Dim ID As Double
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    'Federica dicembre 2017 - Compongo l'inizio delle Tag WinCC solo una volta
    InizioTag = CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000")
    
    'Ultima media oraria
    If DatiPerWinCC Then
        'luca marzo 2017
        Call DefinisciTagOraSemiora
        
        'luca marzo 2017
        Call ElaboraAggiornaUltimaMedia(iIdx1, MedieOra(periodo, iIdx1, 1, numMedia), 0)
        
        '******************** VALIDITA' E COLORAZIONE IN BASE A VALIDITA'
        'luca marzo 2017
        ScriviTag InizioTag & TagValiditaUltimaMedia, StsMedieOra(periodo, iIdx1, 1, numMedia)
        
        If StsMedieOra(periodo, iIdx1, 1, numMedia) = "VAL" Or StsMedieOra(periodo, iIdx1, 1, numMedia) = "AUX" Then
            UltimaMediaOraria = MedieOra(periodo, iIdx1, 1, numMedia)
            'luca marzo 2017
            ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 0
        Else
            UltimaMediaOraria = -9999
            'luca marzo 2017
            ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 2
        End If
        
        '******************** ID
        'luca marzo 2017
        ID = ContaOraOK(periodo, iIdx1, 1, numMedia) / MaxDati * 100
        'controllo che non sia > 100 per sicurezza
        If ID > 100 Then ID = 100
        
        'luca marzo 2017
        ScriviTag InizioTag & TagIDUltimaMedia, ID
        
    End If
    
    'Altre medie da database
    NewDataObj rsDati
    Set Windas = CreateObject("AttimoFwk.Windas")
    
    With rsDati
        'ultima media giornaliera
        Call ElaboraAggiornaMediaGiorno(rsDati, Windas, iIdx1, Format(DateAdd("d", -1, Elabdate), "yyyymmdd"), False)
        'media giornaliera in corso e previsionale
        Call ElaboraAggiornaMediaGiorno(rsDati, Windas, iIdx1, Format(Elabdate, "yyyymmdd"), True)
        
        'ultima media mensile
        Call ElaboraAggiornaMediaMese(rsDati, Windas, iIdx1, DateAdd("m", -1, Elabdate), False)
        'media mensile in corso e previsionale
        Call ElaboraAggiornaMediaMese(rsDati, Windas, iIdx1, Elabdate, True)
        
        If AbilitaTrimestre Then
            'luca luglio 2017
            'ultima media trimestrale
            Call ElaboraAggiornaMediaTrimestre(rsDati, iIdx1, DateAdd("m", -3, DataInizioTrimestre(month(Elabdate))), DateAdd("m", -3, DataFineTrimestre(month(Elabdate))), False)
            'media trimestrale in corso e previsionale
            Call ElaboraAggiornaMediaTrimestre(rsDati, iIdx1, DataInizioTrimestre(month(Elabdate)), DataFineTrimestre(month(Elabdate)), True)
        End If
                
        'Federica febbraio 2018
        'ultima media annuale
        Call ElaboraAggiornaMediaAnno(rsDati, Windas, iIdx1, DateAdd("yyyy", -1, Elabdate), False)
        'media annuale in corso e previsionale
        Call ElaboraAggiornaMediaAnno(rsDati, Windas, iIdx1, Elabdate, False)
        
        'luca marzo 2017
        'If DatiPerWinCC Then
        If DatiPerWinCC And UCase(Tabella) = "WDS_ELAB" Then
            If Abilita48H Then
                Call DefinisciTag48H 'Federica dicembre 2017
            
                'ultima media 48h
                Call ElaboraAggiornaUltimaMedia(iIdx1, UltimaMedia48h, -1)
                'media 48h in corso e previsionale
                Call ElaboraAggiornaMediaCorrente(iIdx1, MediaInCorso48h, -1)
                Call ElaboraAggiornaMediaPrevisionale(iIdx1, ProiezioneMedia48h, -1)
            
                'luca 06/09/2016 scrivo ID ultima media 48H e ID media 48H in costruzione
                ScriviTag InizioTag & TagIDUltimaMedia, IDUltima48H
                ScriviTag InizioTag & TagIDMediaCorrente, IDCostruzione48H
                
                'luca 16/09/2016 scrivo validità media 48H ultima e in costruzione
                ScriviTag InizioTag & TagValiditaUltimaMedia, StatusUltima48H
                ScriviTag InizioTag & TagValiditaMediaCorrente, StatusCostruzione48H
                
                'luca 16/09/2016 gestisco colorazione celle validità media 48H ultima e in costruzione
                ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, IIf(StatusUltima48H = "VAL", 0, 2)
                
                'luca 16/09/2016 gestisco colorazione celle validità media 48H ultima e in costruzione
                ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, IIf(StatusCostruzione48H = "VAL", 0, 2)
            End If
        End If
    End With
        
    Set rsDati = Nothing
    Set Windas = Nothing
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMedie: " + Error(Err), 1)

End Sub

'Federica dicembre 2017
Sub ElaboraAggiornaMediaCorrente(ByVal ii, ByVal valore, ByVal tipo)

    On Error GoTo GestErrore

    Call ElaboraAggiornaWinCCTag(ii, InizioTag & TagMediaCorrente, InizioTag & TagVisualizzazioneMediaCorrente, valore, tipo)
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraAggiornaMediaCorrente: " & Error(Err()), 1)

End Sub

'Federica dicembre 2017
Sub ElaboraAggiornaUltimaMedia(ByVal ii, ByVal valore, ByVal tipo)

    On Error GoTo GestErrore

    Call ElaboraAggiornaWinCCTag(ii, InizioTag & TagUltimaMedia, InizioTag & TagVisualizzazioneUltimaMedia, valore, tipo)
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraAggiornaUltimaMedia: " & Error(Err()), 1)

End Sub

'Federica dicembre 2017
Sub ElaboraAggiornaMediaPrevisionale(ByVal ii, ByVal valore, ByVal tipo)

    On Error GoTo GestErrore

    Call ElaboraAggiornaWinCCTag(ii, InizioTag & TagMediaPrevisionale, InizioTag & TagVisualizzazioneMediaPrevisionale, valore, tipo)
    Exit Sub
    
GestErrore:
    Call WindasLog("ElaboraAggiornaMediaPrevisionale: " & Error(Err()), 1)

End Sub

Private Sub ElaboraAggiornaMediaMese(rsDati, Windas, iIdx1, DataStart, Adesso As Boolean)

    Dim strSQL As String
    Dim Media As Double
    Dim Somma As Double
    Dim SommaTot As Double
    Dim Status As String
    Dim ContaValidiInMarcia As Integer
    Dim ContaInMarcia As Integer
    Dim MeasureCod As String
    Dim ID As Double
    Dim Data As String
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    Data = Format(DataStart, "yyyymm")
    MeasureCod = SuperTrim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)
    
    With rsDati
                
        'determina le ore di marcia del mese
        strSQL = Windas.DatoMese(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "", "", "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaInMarcia = .GetValue("DATO")
        End If
        
        'determina le ore valide del mese in condizione di impianto in marcia
        strSQL = Windas.DatoMese(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'")
        
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        
        If Not .iseof Then
            ContaValidiInMarcia = .GetValue("DATO")
        End If
        
        'luca 06/09/2016 calcolo ID%
        If ContaInMarcia > 0 Then
            ID = ContaValidiInMarcia / ContaInMarcia * 100
            'controllo che non sia > 100 per sicurezza
            If ID > 100 Then ID = 100
        Else
            ID = 0
        End If
        
        'FLUSSO DI MASSA
        SommaTot = CalcolaFM(rsDati, MeasureCod, strInTransitorio, Format(DateAdd("d", -1 * day(DataStart), DataStart) + 1, "yyyymmdd"), Format(DateAdd("m", 1, DateAdd("d", -1 * day(DataStart), DataStart) + 1) - 1, "yyyymmdd"))
        Somma = CalcolaFM(rsDati, MeasureCod, "'30'", Format(DateAdd("d", -1 * day(DataStart), DataStart) + 1, "yyyymmdd"), Format(DateAdd("m", 1, DateAdd("d", -1 * day(DataStart), DataStart) + 1) - 1, "yyyymmdd"))
        
        strSQL = Windas.DatoMese(Tabella, "AVG", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'") & " group by dt_stationcode"
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        
        If Not .iseof Then
            Media = rsDati.GetValue("DATO")
            'luca 16/09/2016
            Status = "VAL"
        Else
            Media = -9999
            'luca 16/09/2016
            Status = "ERR"
        End If

        'aggiorna WinCC
        If DatiPerWinCC Then
            'luca marzo 2017
            Call DefinisciTagMese
            
            'luca marzo 2017
            'Federica ottobre 2017 - Passo Tipo = 2 per gestire le soglie mensili
            If Not Adesso Then
                Call ElaboraAggiornaUltimaMedia(iIdx1, Media, 2)
            Else
                Call ElaboraAggiornaMediaCorrente(iIdx1, Media, 2)
            End If
        End If
        
        'luca 16/09/2016 controllo ID
        If ID < 80 Then
            Status = "ERR"
        End If
        
        'luca marzo 2017
        If Adesso Then
            'Alby Dicembre 2015 media previsionale giornaliera solo se elabora la giornata odierna
            If DatiPerWinCC Then
                
                'luca marzo 2017
                Call ElaboraAggiornaMediaPrevisionale(iIdx1, ProiezioneMediaMensile(Media), -1)
                
                'luca marzo 2017
                ScriviTag InizioTag & TagIDMediaCorrente, ID
                
                'luca marzo 2017
                ScriviTag InizioTag & TagValiditaMediaCorrente, Status
                
                'luca 16/09/2016 colorazione cella validità
                If Status = "VAL" Then
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 0
                Else
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 2
                End If
            End If
        Else
            'luca 11/10/2016 controllo sulle 144 ore di marcia
            'luca marzo 2017
            If UCase(Tabella) = "WDS_ELAB" Then
                If ContaInMarcia < 144 Then
                    Status = "ERR"
                End If
            ElseIf UCase(Tabella) = "WDS_HALF" Then
                If ContaInMarcia < 288 Then
                    Status = "ERR"
                End If
            End If
            
            'luca 06/09/2016 scrivo ID ultimo mese
            If DatiPerWinCC Then
                'luca marzo 2017
                ScriviTag InizioTag & TagIDUltimaMedia, ID
                
                'luca 16/09/2016 validità
                'luca marzo 2017
                ScriviTag InizioTag & TagValiditaUltimaMedia, Status
                
                'luca 16/09/2016 colorazione cella validità
                If Status = "VAL" Then
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 0
                Else
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 2
                End If
            End If
        End If
        
        'luca luglio 2017
        If Not Client Then
            Call ElaboraAggiornaSQL("wds_month", Data, iIdx1, Media, Status, Somma, SommaTot, ContaValidiInMarcia, ContaInMarcia)
        End If
    
    End With

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMediaMese: " + Error(Err), 1)
    Resume Next
    
End Sub

'luca luglio 2017
Private Sub ElaboraAggiornaMediaTrimestre(rsDati, iIdx1, DataInizioTrimestre As Date, DataFineTrimestre As Date, Adesso As Boolean)

    Dim strSQL As String
    Dim Media As Double
    Dim Somma As Double
    Dim SommaTot As Double
    Dim Status As String
    Dim ContaValidiInMarcia As Integer
    Dim ContaInMarcia As Integer
    Dim MeasureCod As String
    Dim ID As Double
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
    MeasureCod = SuperTrim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)
        
    With rsDati
        
        'ORE DI MARCIA
        strSQL = "SELECT COUNT(DT_VALUE) AS DATO FROM " & Tabella & " WHERE DT_STATIONCODE = '" & StationCode & "' AND DT_DATE>= '" & Format(DataInizioTrimestre, "yyyymmdd") & "'"
        strSQL = strSQL & " AND DT_DATE <='" & Format(DataFineTrimestre, "yyyymmdd") & "' AND DT_MEASURECOD = '" & MeasureCod & "' AND DT_CUSTOM1 = '30'"
        
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaInMarcia = .GetValue("DATO")
        End If


        'ORE DI MARCIA VALIDE
        strSQL = "SELECT COUNT(DT_VALUE) AS DATO FROM " & Tabella & " WHERE DT_STATIONCODE = '" & StationCode & "' AND DT_DATE>= '" & Format(DataInizioTrimestre, "yyyymmdd") & "'"
        strSQL = strSQL & " AND DT_DATE <='" & Format(DataFineTrimestre, "yyyymmdd") & "' AND DT_MEASURECOD = '" & MeasureCod & "' AND DT_CUSTOM1 = '30' AND DT_VALIDFLAG IN (" & strValidValidflags & ")"

        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaValidiInMarcia = .GetValue("DATO")
        End If
        
        
        'ID
        If ContaInMarcia > 0 Then
            ID = ContaValidiInMarcia / ContaInMarcia * 100
            'controllo che non sia > 100 per sicurezza
            If ID > 100 Then ID = 100
        Else
            ID = 0
        End If
        
        'FLUSSO DI MASSA
        SommaTot = CalcolaFM(rsDati, MeasureCod, strInTransitorio, DataInizioTrimestre, DataFineTrimestre)
        Somma = CalcolaFM(rsDati, MeasureCod, "'30'", DataInizioTrimestre, DataFineTrimestre)
        
        'MEDIA
        strSQL = "SELECT AVG(DT_VALUE) AS DATO FROM " & Tabella & " WHERE DT_STATIONCODE = '" & StationCode & "' AND DT_DATE>= '" & Format(DataInizioTrimestre, "yyyymmdd") & "'"
        strSQL = strSQL & " AND DT_DATE <='" & Format(DataFineTrimestre, "yyyymmdd") & "' AND DT_MEASURECOD = '" & MeasureCod & "' AND DT_CUSTOM1 = '30' AND DT_VALIDFLAG IN (" & strValidValidflags & ") GROUP BY DT_STATIONCODE"
        
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            Media = .GetValue("DATO")
            Status = "VAL"
        Else
            Media = -9999
            Status = "ERR"
        End If
            
        If DatiPerWinCC Then
            Call DefinisciTagTrimestre
            
            If Not Adesso Then
                'Call ElaboraAggiornaWinCCTag(iIdx1, InizioTag & TagUltimaMedia, InizioTag & TagVisualizzazioneUltimaMedia, Media, -1)
                Call ElaboraAggiornaUltimaMedia(iIdx1, Media, -1)
            Else
                'Call ElaboraAggiornaWinCCTag(iIdx1, InizioTag & TagMediaCorrente, InizioTag & TagVisualizzazioneMediaCorrente, Media, -1)
                Call ElaboraAggiornaMediaCorrente(iIdx1, Media, -1)
            End If
        End If
        
        'PER ORA GESTISCO IL 70%
        If ID < 70 Then
            Status = "ERR"
        End If
        
        If Adesso Then
            If DatiPerWinCC Then
                
                'PROIEZIONE TRIMESTRALE
                'Call ElaboraAggiornaWinCCTag(iIdx1, InizioTag & TagMediaPrevisionale, InizioTag & TagVisualizzazioneMediaPrevisionale, ProiezioneMediaTrimestrale(Media, DataInizioTrimestre, DataFineTrimestre), -1)
                Call ElaboraAggiornaMediaPrevisionale(iIdx1, ProiezioneMediaTrimestrale(Media, DataInizioTrimestre, DataFineTrimestre), -1)
                
                'ID
                ScriviTag InizioTag & TagIDMediaCorrente, ID
                
                'VALIDITA'
                ScriviTag InizioTag & TagValiditaMediaCorrente, Status
                
                'VISUALIZZAZIONE VALIDITA'
                If Status = "VAL" Then
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 0
                Else
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 2
                End If
            End If
        Else
            'PER ORA NON GESTISCO
'            If UCase(Tabella) = "WDS_ELAB" Then
'                If ContaInMarcia < 144 Then
'                    Status = "ERR"
'                End If
'            ElseIf UCase(Tabella) = "WDS_HALF" Then
'                If ContaInMarcia < 288 Then
'                    Status = "ERR"
'                End If
'            End If
            
            If DatiPerWinCC Then
                
                'ID
                ScriviTag InizioTag & TagIDUltimaMedia, ID
                
                'VALIDITA'
                ScriviTag InizioTag & TagValiditaUltimaMedia, Status
                
                'VISUALIZZAZIONE VALIDITA'
                If Status = "VAL" Then
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 0
                Else
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 2
                End If
            End If
        End If
        
        If Not Client Then
            Call ElaboraAggiornaSQL("wds_quarterly", Format(DataInizioTrimestre, "yyyymm"), iIdx1, Media, Status, Somma, SommaTot, ContaValidiInMarcia, ContaInMarcia)
        End If
    
    End With

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMediaTrimestre: " + Error(Err), 1)
    Resume Next
    
End Sub

Private Sub DefinisciTagOraSemiora()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_MONU"
            TagValiditaUltimaMedia = "_MONU_VAL"
            TagVisualizzazioneUltimaMedia = "_MONU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MONU_VAL_VIS"
            TagIDUltimaMedia = "_MONU_ID"
        Case "WDS_HALF"
            TagUltimaMedia = "_MSNU"
            TagValiditaUltimaMedia = "_MSNU_VAL"
            TagVisualizzazioneUltimaMedia = "_MSNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MSNU_VAL_VIS"
            TagIDUltimaMedia = "_MSNU_ID"
    End Select

    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTagOraSemiora: " + Error(Err), 1)

End Sub

'Federica dicembre 2017
Private Sub DefinisciTag48H()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_M48HNU"
            TagValiditaUltimaMedia = "_M48HNU_VAL"
            TagVisualizzazioneUltimaMedia = "_M48HNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_M48HNU_VAL_VIS"
            TagIDUltimaMedia = "_M48HNU_ID"
            TagMediaCorrente = "_M48HNC"
            TagValiditaMediaCorrente = "_M48HNC_VAL"
            TagVisualizzazioneMediaCorrente = "_M48HNC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_M48HNC_VAL_VIS"
            TagIDMediaCorrente = "_M48HNC_ID"
            TagMediaPrevisionale = "_M48HNP"
            TagVisualizzazioneMediaPrevisionale = "_M48HNP_VIS"
    End Select
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTag48H: " + Error(Err), 1)

End Sub

Private Sub DefinisciTagAnno()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_MAONU"
            TagValiditaUltimaMedia = "_MAONU_VAL"
            TagVisualizzazioneUltimaMedia = "_MAONU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MAONU_VAL_VIS"
            TagIDUltimaMedia = "_MAONU_ID"
            TagMediaCorrente = "_MAONC"
            TagValiditaMediaCorrente = "_MAONC_VAL"
            TagVisualizzazioneMediaCorrente = "_MAONC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MAONC_VAL_VIS"
            TagIDMediaCorrente = "_MAONC_ID"
            TagMediaPrevisionale = "_MAONP"
            TagVisualizzazioneMediaPrevisionale = "_MAONP_VIS"
        Case "WDS_HALF"
            TagUltimaMedia = "_MASNU"
            TagValiditaUltimaMedia = "_MASNU_VAL"
            TagVisualizzazioneUltimaMedia = "_MASNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MASNU_VAL_VIS"
            TagIDUltimaMedia = "_MASNU_ID"
            TagMediaCorrente = "_MASNC"
            TagValiditaMediaCorrente = "_MASNC_VAL"
            TagVisualizzazioneMediaCorrente = "_MASNC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MASNC_VAL_VIS"
            TagIDMediaCorrente = "_MASNC_ID"
            TagMediaPrevisionale = "_MASNP"
            TagVisualizzazioneMediaPrevisionale = "_MASNP_VIS"
    End Select
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTagMese: " + Error(Err), 1)

End Sub

Private Sub ElaboraAggiornaMediaGiorno(rsDati, Windas, iIdx1, Data As String, Adesso As Boolean)

    Dim strSQL As String
    Dim Media As Double
    Dim Somma As Double
    Dim SommaTot As Double  'Federica dicembre 2017
    Dim Status As String
    Dim ContaValidiInMarcia As Integer
    Dim ContaInMarcia As Integer
    Dim MeasureCod As String
    Dim ID As Double
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    MeasureCod = SuperTrim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)
    
    With rsDati
            
        'determina le ore di marcia della giornata
        'luca marzo 2017
        strSQL = Windas.DatoGiorno(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "", "", "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaInMarcia = .GetValue("DATO")
        End If
            
        'determina le ore valide della giornata in condizione di impianto in marcia
        'luca marzo 2017
        strSQL = Windas.DatoGiorno(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaValidiInMarcia = .GetValue("DATO")
        End If
        
        'luca 06/09/2016 calcolo ID%
        If ContaInMarcia > 0 Then
            ID = ContaValidiInMarcia / ContaInMarcia * 100
            'controllo che non sia > 100 per sicurezza
            If ID > 100 Then ID = 100
        Else
            ID = 0
        End If
        
        'FLUSSO DI MASSA
        SommaTot = CalcolaFM(rsDati, MeasureCod, strInTransitorio, Data, Data)
        Somma = CalcolaFM(rsDati, MeasureCod, "'30'", Data, Data)
                
        'determina la media delle ore valide in condizione di impianto in marcia
        'luca marzo 2017
        strSQL = Windas.DatoGiorno(Tabella, "AVG", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            Media = .GetValue("DATO")
            Status = "VAL"
        Else
            Media = -9999
            Status = "ERR"
        End If

        'aggiorna WinCC
        If DatiPerWinCC Then
            'luca marzo 2017
            Call DefinisciTagGiorno
            'luca marzo 2017
            If Not Adesso Then
                Call ElaboraAggiornaUltimaMedia(iIdx1, Media, 1)
            Else
                Call ElaboraAggiornaMediaCorrente(iIdx1, Media, 1)
            End If
        End If
        
        'luca 16/09/2016 controllo ID (media giornaliera 70%)
        If ID < 70 Then
            Status = "ERR"
        End If
        
        'luca marzo 2017
        If Adesso Then
            'Alby Dicembre 2015 media previsionale giornaliera solo se elabora la giornata odierna
            UltimaMediaGiorno = Media
            
            If DatiPerWinCC Then
                
                'luca marzo 2017
                Call ElaboraAggiornaMediaPrevisionale(iIdx1, ProiezioneMediaGiornaliera(Media), 1)
                'luca marzo 2017
                ScriviTag InizioTag & TagIDMediaCorrente, ID
                
                'luca 16/09/2016 validità
                ScriviTag InizioTag & TagValiditaMediaCorrente, Status
                
                'luca 16/09/2016 colorazione cella validità
                If Status = "VAL" Then
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 0
                Else
                    ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, 2
                End If
            End If
        Else
            'luca marzo 2017
            If UCase(Tabella) = "WDS_ELAB" Then
                If ContaInMarcia < 6 Then
                    Status = "ERR"
                End If
            ElseIf UCase(Tabella) = "WDS_HALF" Then
                 If ContaInMarcia < 12 Then
                    Status = "ERR"
                End If
            End If
            
            'luca 06/09/2016 scrivo ID ultimo giorno
            If DatiPerWinCC Then
                'luca marzo 2017
                ScriviTag InizioTag & TagIDUltimaMedia, ID
                
                'luca marzo 2017
                ScriviTag InizioTag & TagValiditaUltimaMedia, Status
                
                'luca 16/09/2016 colorazione cella validità
                If Status = "VAL" Then
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 0
                Else
                    'luca marzo 2017
                    ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, 2
                End If
            End If
        End If
        
        'luca luglio 2017
        If Not Client Then
            Call ElaboraAggiornaSQL("wds_days", Data, iIdx1, Media, Status, Somma, SommaTot, ContaValidiInMarcia, ContaInMarcia)
        End If
    
    End With

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMediaGiorno: " + Error(Err), 1)
    Resume Next
    
End Sub

Private Sub ElaboraAggiornaMediaAnno(rsDati, Windas, iIdx1, DataStart, Adesso As Boolean)

    Dim strSQL As String
    Dim Media As Double
    Dim Somma As Double
    Dim SommaTot As Double
    Dim Status As String
    Dim ContaValidiInMarcia As Integer
    Dim ContaInMarcia As Integer
    Dim MeasureCod As String
    Dim ID As Double
    
    'Federica dicembre 2017
    Dim Data As String  'Per compatibilita
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
        
    'Federica dicembre 2017
    Data = Format(DataStart, "yyyy")
    MeasureCod = SuperTrim(gaConfigurazioneArchivio(iIdx1).STRUM.NomeParametro)
    
    With rsDati
                
        'determina le ore di marcia dell'anno
        strSQL = Windas.DatoMese(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "", "", "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaInMarcia = .GetValue("DATO")
        End If
        
        'determina le ore valide dell'anno in condizione di impianto in marcia
        strSQL = Windas.DatoMese(Tabella, "COUNT", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'")
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        If Not .iseof Then
            ContaValidiInMarcia = .GetValue("DATO")
        End If
        
        'luca 06/09/2016 calcolo ID%
        If ContaInMarcia > 0 Then
            ID = ContaValidiInMarcia / ContaInMarcia * 100
            'controllo che non sia > 100 per sicurezza
            If ID > 100 Then ID = 100
        Else
            ID = 0
        End If
        
        'FLUSSO DI MASSA
        SommaTot = CalcolaFM(rsDati, MeasureCod, strInTransitorio, CStr(year(DataStart)) & "0101", CStr(year(DataStart)) & "1231")
        Somma = CalcolaFM(rsDati, MeasureCod, "'30'", CStr(year(DataStart)) & "0101", CStr(year(DataStart)) & "1231")
        
        strSQL = Windas.DatoMese(Tabella, "AVG", StationCode, Data, MeasureCod, "DT_VALUE", "DT_VALIDFLAG", strValidValidflags, "'30'") & " group by dt_stationcode"
        If strSQL <> "" Then .ExecuteSQL (strSQL)
        
        If Not .iseof Then
            Media = rsDati.GetValue("DATO")
            'luca 16/09/2016
            Status = "VAL"
        Else
            Media = -9999
            'luca 16/09/2016
            Status = "ERR"
        End If

        'aggiorna WinCC
        If DatiPerWinCC Then
            'luca marzo 2017
            Call DefinisciTagAnno
            'Federica ottobre 2017 - Passo Tipo = 2 per gestire le soglie mensili
            If Adesso Then
                Call ElaboraAggiornaMediaCorrente(iIdx1, Media, 2)
            Else
                Call ElaboraAggiornaUltimaMedia(iIdx1, Media, 2)
            End If
        End If
        
        'luca 16/09/2016 controllo ID
        If ID = 0 Then
            Status = "ERR"
        End If
        
        'luca marzo 2017
        If Adesso Then
            'Alby Dicembre 2015 media previsionale giornaliera solo se elabora la giornata odierna
            If DatiPerWinCC Then
                Call ElaboraAggiornaMediaPrevisionale(iIdx1, ProiezioneMediaMensile(Media), -1)
                
                ScriviTag InizioTag & TagIDMediaCorrente, ID
                ScriviTag InizioTag & TagValiditaMediaCorrente, Status
                
                'luca 16/09/2016 colorazione cella validità
                ScriviTag InizioTag & TagVisualizzazioneValiditaMediaCorrente, IIf(Status = "VAL", 0, 2)
            End If
        Else
            'luca 06/09/2016 scrivo ID ultimo mese
            If DatiPerWinCC Then
                ScriviTag InizioTag & TagIDUltimaMedia, ID
                ScriviTag InizioTag & TagValiditaUltimaMedia, Status
                
                'luca 16/09/2016 colorazione cella validità
                ScriviTag InizioTag & TagVisualizzazioneValiditaUltimaMedia, IIf(Status = "VAL", 0, 2)
            End If
        End If
        
        'luca luglio 2017
        If Not Client Then
            Call ElaboraAggiornaSQL("wds_year", Data, iIdx1, Media, Status, Somma, SommaTot, ContaValidiInMarcia, ContaInMarcia)
        End If
    
    End With

    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMediaAnno: " + Error(Err), 1)
    Resume Next
    
End Sub

Private Sub DefinisciTagGiorno()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_MGONU"
            TagValiditaUltimaMedia = "_MGONU_VAL"
            TagVisualizzazioneUltimaMedia = "_MGONU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MGONU_VAL_VIS"
            TagIDUltimaMedia = "_MGONU_ID"
            TagMediaCorrente = "_MGONC"
            TagValiditaMediaCorrente = "_MGONC_VAL"
            TagVisualizzazioneMediaCorrente = "_MGONC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MGONC_VAL_VIS"
            TagIDMediaCorrente = "_MGONC_ID"
            TagMediaPrevisionale = "_MGONP"
            TagVisualizzazioneMediaPrevisionale = "_MGONP_VIS"
        Case "WDS_HALF"
            TagUltimaMedia = "_MGSNU"
            TagValiditaUltimaMedia = "_MGSNU_VAL"
            TagVisualizzazioneUltimaMedia = "_MGSNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MGSNU_VAL_VIS"
            TagIDUltimaMedia = "_MGSNU_ID"
            TagMediaCorrente = "_MGSNC"
            TagValiditaMediaCorrente = "_MGSNC_VAL"
            TagVisualizzazioneMediaCorrente = "_MGSNC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MGSNC_VAL_VIS"
            TagIDMediaCorrente = "_MGSNC_ID"
            TagMediaPrevisionale = "_MGSNP"
            TagVisualizzazioneMediaPrevisionale = "_MGSNP_VIS"
    End Select
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTagGiorno: " + Error(Err), 1)

End Sub

Private Sub DefinisciTagMese()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_MMONU"
            TagValiditaUltimaMedia = "_MMONU_VAL"
            TagVisualizzazioneUltimaMedia = "_MMONU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MMONU_VAL_VIS"
            TagIDUltimaMedia = "_MMONU_ID"
            TagMediaCorrente = "_MMONC"
            TagValiditaMediaCorrente = "_MMONC_VAL"
            TagVisualizzazioneMediaCorrente = "_MMONC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MMONC_VAL_VIS"
            TagIDMediaCorrente = "_MMONC_ID"
            TagMediaPrevisionale = "_MMONP"
            TagVisualizzazioneMediaPrevisionale = "_MMONP_VIS"
        Case "WDS_HALF"
            TagUltimaMedia = "_MMSNU"
            TagValiditaUltimaMedia = "_MMSNU_VAL"
            TagVisualizzazioneUltimaMedia = "_MMSNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MMSNU_VAL_VIS"
            TagIDUltimaMedia = "_MMSNU_ID"
            TagMediaCorrente = "_MMSNC"
            TagValiditaMediaCorrente = "_MMSNC_VAL"
            TagVisualizzazioneMediaCorrente = "_MMSNC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MMSNC_VAL_VIS"
            TagIDMediaCorrente = "_MMSNC_ID"
            TagMediaPrevisionale = "_MMSNP"
            TagVisualizzazioneMediaPrevisionale = "_MMSNP_VIS"
    End Select
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTagMese: " + Error(Err), 1)

End Sub

Private Sub DefinisciTagTrimestre()
    
    Select Case UCase(Tabella)
        Case "WDS_ELAB"
            TagUltimaMedia = "_MTONU"
            TagValiditaUltimaMedia = "_MTONU_VAL"
            TagVisualizzazioneUltimaMedia = "_MTONU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MTONU_VAL_VIS"
            TagIDUltimaMedia = "_MTONU_ID"
            TagMediaCorrente = "_MTONC"
            TagValiditaMediaCorrente = "_MTONC_VAL"
            TagVisualizzazioneMediaCorrente = "_MTONC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MTONC_VAL_VIS"
            TagIDMediaCorrente = "_MTONC_ID"
            TagMediaPrevisionale = "_MTONP"
            TagVisualizzazioneMediaPrevisionale = "_MTONP_VIS"
        Case "WDS_HALF"
            TagUltimaMedia = "_MTSNU"
            TagValiditaUltimaMedia = "_MTSNU_VAL"
            TagVisualizzazioneUltimaMedia = "_MTSNU_VIS"
            TagVisualizzazioneValiditaUltimaMedia = "_MTSNU_VAL_VIS"
            TagIDUltimaMedia = "_MTSNU_ID"
            TagMediaCorrente = "_MTSNC"
            TagValiditaMediaCorrente = "_MTSNC_VAL"
            TagVisualizzazioneMediaCorrente = "_MTSNC_VIS"
            TagVisualizzazioneValiditaMediaCorrente = "_MTSNC_VAL_VIS"
            TagIDMediaCorrente = "_MTSNC_ID"
            TagMediaPrevisionale = "_MTSNP"
            TagVisualizzazioneMediaPrevisionale = "_MTSNP_VIS"
    End Select
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DefinisciTagTrimestre: " + Error(Err), 1)

End Sub

Sub ElaboraAggiornaMedie10minuti(periodo As Integer, iIdx1 As Integer, numMedia As Integer)

    Dim ID As Double
    
    'Alby Dicembre 2015
    On Error GoTo GestErrore
    
    'Ultima media oraria
    If DatiPerWinCC Then
       
        Call ElaboraAggiornaWinCCTag(iIdx1, CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU", CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU_VIS", MedieOra(periodo, iIdx1, 1, numMedia), 0)
        
        ScriviTag CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU_VAL", StsMedieOra(periodo, iIdx1, 1, numMedia)
        
        If StsMedieOra(periodo, iIdx1, 1, numMedia) = "VAL" Or StsMedieOra(periodo, iIdx1, 1, numMedia) = "AUX" Then
            ScriviTag CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU_VAL_VIS", 0
        Else
            ScriviTag CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU_VAL_VIS", 2
        End If
        
        ID = ContaOraOK(periodo, iIdx1, 1, numMedia) / MaxDati * 100
        'controllo che non sia > 100 per sicurezza
        If ID > 100 Then ID = 100

        ScriviTag CStr(NumeroLineaBFData) & ".AM" & Format(gaConfigurazioneArchivio(iIdx1).STRUM.CodiceParametro, "000") & "_M10MNU_ID", ID
        
    End If
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata ElaboraAggiornaMedie10minuti: " + Error(Err), 1)

End Sub
