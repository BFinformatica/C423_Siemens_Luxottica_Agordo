Attribute VB_Name = "QAL3"
Option Explicit

'***** QAL3 *****
Public zero(20)
Public span(20)
Public ParamSelected(20) As Boolean 'Nicolò Agosto 2016

Sub ControlloTarature(ByVal stringaMisure As String)
'TODO: Nella parametrizzazione inserire stringa con le misure coinvolte (es. 0;2;8)

    Dim Misure() As String
    Dim i As Integer
    
    On Error GoTo Gesterrore
    
    If stringaMisure = "" Then Exit Sub
    
    'aggiunta misure
    Misure = Split(stringaMisure, ";")
'    Misure(0) = 0 'CO
 '   Misure(1) = 2 'O2
  '  Misure(2) = 8 'NO
    
    For i = 0 To UBound(Misure)
        Call GestioneResetQAL3(CInt(Misure(i)))
    Next i
        
    'Alby Febbraio 2016                 Sequenza, QAL3incorso, QAL3finitaOK
    Call LeggiVariabiliWinCCperTarature(1, 44, 45)

    Exit Sub
    
Gesterrore:
    Call WindasLog("ControlloTarature " + Error(Err), 1, OPC)

End Sub

Private Sub LeggiVariabiliWinCCperTarature(Sequenza, QAL3inCorso, QAL3ultimata)
'TODO: Eventualmente configurare le Tag se la prevedono

    Static statoQAL3(2) As Integer
    Static cont(2) As Integer
    'luca aprile 2017
    Dim i As Integer
    Dim tempQAL3(9) As Double
    Dim tempQAL3Selettori(9) As Boolean
    
    'leggo sempre risultati QAL3 e selettori
    'luca maggio 2018 disabilito no OPC
'    For i = 0 To 5
'        tempQAL3(i) = CDec(LeggiTag(CStr(NumeroLinea) & "_DB80_QAL3_" & Format(i, "00")))
'        'tempQAL3Selettori(i) = CBool(LeggiTag(CStr(NumeroLinea) & "_MSKCAL_" & Format(i, "00")))
'    Next i
    
    On Error GoTo Gesterrore
    
    If Valore_DI(QAL3inCorso) = 1 Then
        statoQAL3(Sequenza) = 1
    End If
    
    If statoQAL3(Sequenza) = 1 Then
        If Valore_DI(QAL3ultimata) = 1 Then
            'luca 22/09/2016 inserisco contatore perchè salva i risultati troppo presto (nuovi risultati non ancora a disposizione lato PLC)
            If cont(Sequenza) = 0 Then
            
                Call WindasLog("QAL3 terminata regolarmente...  salvataggio risultati", 0, OPC)
                
                'CO
                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(0), span(0), IIf(NumeroLinea = 1, "AI20", "AI26"), IIf(NumeroLinea = 1, "AI21", "AI27"))
                
                ParamSelected(0) = CBool(LeggiTag("DI136")) And CBool(LeggiTag("DI144"))  'Nicolò Agosto 2016

                Call WindasLog("Parametro CO selettori: " & ParamSelected(0), 0, OPC)
                        
                'NOx
                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(1), span(1), IIf(NumeroLinea = 1, "AI22", "AI28"), IIf(NumeroLinea = 1, "AI23", "AI29"))
                
                ParamSelected(1) = CBool(LeggiTag("DI137")) And CBool(LeggiTag("DI145"))   'Nicolò Agosto 2016
                'ParamSelected(0) = True
                Call WindasLog("Parametro NOX selettori: " & ParamSelected(1), 0, OPC)
                        
                'O2
                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(2), span(2), IIf(NumeroLinea = 1, "AI24", "AI30"), IIf(NumeroLinea = 1, "AI25", "AI31"))
                
                ParamSelected(2) = CBool(LeggiTag("DI138")) And CBool(LeggiTag("DI146"))   'Nicolò Agosto 2016
                'ParamSelected(2) = True
                Call WindasLog("Parametro O2 selettori: " & ParamSelected(2), 0, OPC)
                        
                Call TaratureSalvaQAL3(0, "Q")
                Call TaratureSalvaQAL3(1, "Q")
                Call TaratureSalvaQAL3(2, "Q")
                        
                statoQAL3(Sequenza) = 0
                    
                cont(Sequenza) = 60
            Else
                cont(Sequenza) = cont(Sequenza) - 1
            End If
        End If
    Else
        cont(Sequenza) = 60
    End If
            
    Exit Sub
    
Gesterrore:
    Call WindasLog("LeggiVariabiliWinCCperTarature Sequenza: " + Format(Sequenza, "0") + " " + Error(Err), 1, OPC)

End Sub

'luca aprile 2017
'Private Sub LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(ByRef zero, ByRef span, tagZero As String, tagSpan As String)
Public Sub LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(ByRef zero, ByRef span, tagZero As String, tagSpan As String)
    'Nicolò gennaio 2017
    On Error GoTo Gesterrore
    
    Dim tmpZero As Double 'Non funziona la gestione con i double (999999,9 diventa 999999,875). Uso decimal che dovrebbe essere più preciso.
    Dim tmpSpan As Double
    Dim valuesOk As Boolean
    Dim tryCount As Integer
    
    Const MAX_TRY = 3
    tmpZero = CDec(999999.9)
    tmpSpan = CDec(999999.9)
    valuesOk = False
    tryCount = 0
    While tryCount < MAX_TRY And Not valuesOk
        tmpZero = CDec(LeggiTag(tagZero))
        tmpSpan = CDec(LeggiTag(tagSpan))
        valuesOk = Not (tmpZero = tmpSpan Or tmpZero = CDec(999999.9) Or tmpSpan = CDec(999999.9) Or tmpZero = CDec(0) Or tmpSpan = CDec(0))
        tryCount = tryCount + 1
    Wend
    If tryCount >= MAX_TRY And Not valuesOk Then Call WindasLog("Possibile fallito caricamento tag " & tagZero & " (" & tmpZero & ") e/o tag " & tagSpan & " (" & tmpSpan & ")", 0, OPC)
    zero = tmpZero
    span = tmpSpan
       
    Exit Sub
Gesterrore:
    Call WindasLog("LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC : " & Error(Err), 1, OPC)
    Resume Next
End Sub

'Private Sub TaratureSalvaQAL3(ByVal vnIndiceParametro, ByVal TipoCal)
Public Sub TaratureSalvaQAL3(ByVal vnIndiceParametro, ByVal TipoCal)

    Dim nCanaleLibero
    Dim sOldMsg
    Dim rsQAL3 As Object
    Dim strSQL
    Dim OffsetDrift
    Dim AmplDrift
    Dim DeltaOffsetDrift
    Dim DeltaAmplDrift
    Dim ZeroRif
    Dim SpanRif
    Dim ZeroErr
    Dim SpanErr
    Dim ErroreStr
    Dim nFile
    Dim ndx, OraTaratura, nn
    
    'Alby Dicembre 2015 QAL3
    Dim ZeroResult As Double
    Dim SpanResult As Double
    
    On Error GoTo Gesterrore
    
    NewDataObj rsQAL3
    
    'Nicolò Agosto 2016
    If Not ParamSelected(vnIndiceParametro) Then
        Call WindasLog("Verifica di QAL3 non selezionata per parametro:" + ParametriStrumenti(vnIndiceParametro).NomeParametro, 0, OPC)
        Exit Sub
    End If
    
    ZeroRif = ParametriStrumenti(vnIndiceParametro).ZeroTeorico
    SpanRif = ParametriStrumenti(vnIndiceParametro).SpanTeorico
    ZeroResult = zero((vnIndiceParametro))
    SpanResult = span((vnIndiceParametro))
    
    If ZeroResult >= 999999 Or SpanResult >= 999999 Then
        Call WindasLog("Verifica di QAL3 esclusa per parametro:" + ParametriStrumenti(vnIndiceParametro).NomeParametro, 0, OPC)
        Exit Sub
    End If
    
    '********************** simulazione valori per test **************************
'   ZeroResult = 9
'   SpanResult = 416
'   ZeroRif = 0
'   SpanRif  = 400
    
    '**************************************************************************
    '*                      verifica di PRECISIONE                            *
    '**************************************************************************

    Dim ZERO_hs
    Dim SPAN_hs
    Dim ZERO_ks
    Dim SPAN_ks
    
    Dim OLD_ZERO_dt
    Dim OLD_SPAN_dt
    Dim OLD_ZERO_st
    Dim OLD_SPAN_st
    
    Dim ZERO_dt
    Dim SPAN_dt
    Dim ZERO_st
    Dim SPAN_st
    
    Dim ZERO_sp
    Dim SPAN_sp
    
    Dim OLD_ZERO_Nst
    Dim OLD_SPAN_Nst
    Dim ZERO_Nst
    Dim SPAN_Nst
    
    Dim ZERO_check
    Dim SPAN_check
        
    '******** Inizializza scarto tipo ******
    ZERO_hs = 6.9 * (ParametriStrumenti(vnIndiceParametro).ZeroSams) ^ 2       'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.ZERO_S_AMS
    SPAN_hs = 6.9 * (ParametriStrumenti(vnIndiceParametro).SpanSams) ^ 2       'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.SPAN_S_AMS
    ZERO_ks = 1.85 * (ParametriStrumenti(vnIndiceParametro).ZeroSams) ^ 2  'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.ZERO_S_AMS
    SPAN_ks = 1.85 * (ParametriStrumenti(vnIndiceParametro).SpanSams) ^ 2  'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.SPAN_S_AMS
    
    If ZERO_hs = 0 Then ZERO_hs = 1
    If SPAN_hs = 0 Then SPAN_hs = 1
    If ZERO_ks = 0 Then ZERO_ks = 1
    If SPAN_ks = 0 Then SPAN_ks = 1
    
    '********** estrae ultimi dati *********
    strSQL = "SELECT * FROM WDS_CALIBRATION WHERE CL_STATION = '" & CStr(gsClienteDi) & "' AND CL_PARAMETER = '" & ParametriStrumenti(vnIndiceParametro).NomeParametro & "' ORDER BY cl_date DESC, cl_hour DESC"
    rsQAL3.selectionfast (strSQL)
    If rsQAL3.iseof Then
        '******* inizializza valori *******
        OLD_ZERO_st = 0
        OLD_SPAN_st = 0
        OLD_ZERO_Nst = 0
        OLD_SPAN_Nst = 0
        OLD_ZERO_dt = 0
        OLD_SPAN_dt = 0
    Else
        If rsQAL3.getValue("C9") = "" Then
            OLD_ZERO_st = 0
        Else
            OLD_ZERO_st = Replace(CStr(rsQAL3.getValue("C9")), ".", ",")
        End If
        
        If rsQAL3.getValue("C10") = "" Then
            OLD_SPAN_st = 0
        Else
            OLD_SPAN_st = Replace(CStr(rsQAL3.getValue("C10")), ".", ",")
        End If
        
        If rsQAL3.getValue("C11") = "" Then
            OLD_ZERO_Nst = 0
        Else
            OLD_ZERO_Nst = Replace(CStr(rsQAL3.getValue("C11")), ".", ",")
        End If
        
        If rsQAL3.getValue("C12") = "" Then
            OLD_SPAN_Nst = 0
        Else
             OLD_SPAN_Nst = Replace(CStr(rsQAL3.getValue("C12")), ".", ",")
        End If
        
        If rsQAL3.getValue("C5") = "" Then
            OLD_ZERO_dt = 0
        Else
            OLD_ZERO_dt = Replace(CStr(rsQAL3.getValue("C5")), ".", ",")
        End If
        
        If rsQAL3.getValue("C6") = "" Then
            OLD_SPAN_dt = 0
        Else
            OLD_SPAN_dt = Replace(CStr(rsQAL3.getValue("C6")), ".", ",")
        End If
    End If
    
    '******** verifica di ZERO *******
    ZERO_dt = ZeroResult - ZeroRif
    ZERO_sp = OLD_ZERO_st + ((ZERO_dt - OLD_ZERO_dt) ^ 2) / 2 - ZERO_ks
   
    If ZERO_sp > 0 Then
        ZERO_st = ZERO_sp
        ZERO_Nst = OLD_ZERO_Nst + 1
    Else
        ZERO_st = 0
        ZERO_Nst = 0
    End If
    
    ZERO_check = (ZERO_st >= ZERO_hs)
    
    '******** verifica di SPAN *******
    SPAN_dt = SpanResult - SpanRif
    SPAN_sp = OLD_SPAN_st + ((SPAN_dt - OLD_SPAN_dt) ^ 2) / 2 - SPAN_ks
    
    If SPAN_sp > 0 Then
        SPAN_st = SPAN_sp
        SPAN_Nst = OLD_SPAN_Nst + 1
    Else
        SPAN_st = 0
        SPAN_Nst = 0
    End If
    
    SPAN_check = (SPAN_st >= SPAN_hs)
    
    '**************************************************************************
    '*                     fine verifica di PRECISIONE                        *
    '**************************************************************************
    
    '**************************************************************************
    '*                          verifica di DERIVA                            *
    '**************************************************************************
    Dim ZERO_hx
    Dim SPAN_hx
    Dim ZERO_kx
    Dim SPAN_kx
    
    Dim OLD_ZERO_SUM_POS
    Dim OLD_ZERO_SUM_NEG
    Dim OLD_ZERO_N_POS
    Dim OLD_ZERO_N_NEG
    
    Dim OLD_SPAN_SUM_POS
    Dim OLD_SPAN_SUM_NEG
    Dim OLD_SPAN_N_POS
    Dim OLD_SPAN_N_NEG
    
    Dim ZERO_SUM_POS_p
    Dim ZERO_SUM_NEG_p
    Dim ZERO_SUM_POS_t
    Dim ZERO_SUM_NEG_t
    Dim ZERO_N_POS
    Dim ZERO_N_NEG
    
    Dim SPAN_SUM_POS_p
    Dim SPAN_SUM_NEG_p
    Dim SPAN_SUM_POS_t
    Dim SPAN_SUM_NEG_t
    Dim SPAN_N_POS
    Dim SPAN_N_NEG
    
    Dim ZERO_D_POS_check
    Dim ZERO_D_NEG_check
    Dim SPAN_D_POS_check
    Dim SPAN_D_NEG_check

    Dim ZERO_D_POS
    Dim ZERO_D_NEG
    Dim SPAN_D_POS
    Dim SPAN_D_NEG
    
    Dim ESITO_ZERO
    Dim ESITO_SPAN
    
    '******** Inizializza scarto tipo ******
    ZERO_hx = 2.85 * ParametriStrumenti(vnIndiceParametro).ZeroSams        'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.ZERO_S_AMS
    SPAN_hx = 2.85 * ParametriStrumenti(vnIndiceParametro).SpanSams        'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.SPAN_S_AMS
    ZERO_kx = 0.501 * ParametriStrumenti(vnIndiceParametro).ZeroSams       'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.ZERO_S_AMS
    SPAN_kx = 0.501 * ParametriStrumenti(vnIndiceParametro).SpanSams       'gaConfigurazioneArchivio(vnIndiceParametro).STRUM.SPAN_S_AMS
    
    If ZERO_hx = 0 Then ZERO_hx = 1
    If SPAN_hx = 0 Then SPAN_hx = 1
    If ZERO_kx = 0 Then ZERO_kx = 1
    If SPAN_kx = 0 Then SPAN_kx = 1
    
    If rsQAL3.iseof Then
        '******* inizializza valori *******
        OLD_ZERO_SUM_POS = 0
        OLD_ZERO_SUM_NEG = 0
        OLD_SPAN_SUM_POS = 0
        OLD_SPAN_SUM_NEG = 0
        OLD_ZERO_N_POS = 0
        OLD_ZERO_N_NEG = 0
        OLD_SPAN_N_POS = 0
        OLD_SPAN_N_NEG = 0
        OLD_ZERO_dt = 0
        OLD_SPAN_dt = 0
    Else
        If rsQAL3.getValue("C27") = "" Then
            OLD_ZERO_SUM_POS = 0
        Else
            OLD_ZERO_SUM_POS = Replace(CStr(rsQAL3.getValue("C27")), ".", ",")
        End If
        
        If rsQAL3.getValue("C30") = "" Then
            OLD_ZERO_SUM_NEG = 0
        Else
            OLD_ZERO_SUM_NEG = Replace(CStr(rsQAL3.getValue("C30")), ".", ",")
        End If
        
        If rsQAL3.getValue("C33") = "" Then
            OLD_SPAN_SUM_POS = 0
        Else
            OLD_SPAN_SUM_POS = Replace(CStr(rsQAL3.getValue("C33")), ".", ",")
        End If
        
        If rsQAL3.getValue("C36") = "" Then
            OLD_SPAN_SUM_NEG = 0
        Else
            OLD_SPAN_SUM_NEG = Replace(CStr(rsQAL3.getValue("C36")), ".", ",")
        End If
        
        If rsQAL3.getValue("C28") = "" Then
            OLD_ZERO_N_POS = 0
        Else
            OLD_ZERO_N_POS = Replace(CStr(rsQAL3.getValue("C28")), ".", ",")
        End If
        
        If rsQAL3.getValue("C31") = "" Then
            OLD_ZERO_N_NEG = 0
        Else
            OLD_ZERO_N_NEG = Replace(CStr(rsQAL3.getValue("C31")), ".", ",")
        End If
        
        If rsQAL3.getValue("C34") = "" Then
            OLD_SPAN_N_POS = 0
        Else
            OLD_SPAN_N_POS = Replace(CStr(rsQAL3.getValue("C34")), ".", ",")
        End If
        
        If rsQAL3.getValue("C37") = "" Then
            OLD_SPAN_N_NEG = 0
        Else
            OLD_SPAN_N_NEG = Replace(CStr(rsQAL3.getValue("C37")), ".", ",")
        End If
    
        If rsQAL3.getValue("C5") = "" Then
            OLD_ZERO_dt = 0
        Else
            OLD_ZERO_dt = Replace(CStr(rsQAL3.getValue("C5")), ".", ",")
        End If
        
        If rsQAL3.getValue("C6") = "" Then
            OLD_SPAN_dt = 0
        Else
            OLD_SPAN_dt = Replace(CStr(rsQAL3.getValue("C6")), ".", ",")
        End If
    End If
    
    '******** verifica di ZERO *******
    ZERO_dt = ZeroResult - ZeroRif
    ZERO_SUM_POS_p = OLD_ZERO_SUM_POS + ZERO_dt - ZERO_kx
    ZERO_SUM_NEG_p = OLD_ZERO_SUM_NEG - ZERO_dt - ZERO_kx
    
    If ZERO_SUM_POS_p > 0 Then
        ZERO_SUM_POS_t = ZERO_SUM_POS_p
        ZERO_N_POS = OLD_ZERO_N_POS + 1
    Else
        ZERO_SUM_POS_t = 0
        ZERO_N_POS = 0
    End If
    
    If ZERO_SUM_NEG_p > 0 Then
        ZERO_SUM_NEG_t = ZERO_SUM_NEG_p
        ZERO_N_NEG = OLD_ZERO_N_NEG + 1
    Else
        ZERO_SUM_NEG_t = 0
        ZERO_N_NEG = 0
    End If
    
    ESITO_ZERO = "'NESSUNA DERIVA'"
    ZERO_D_POS_check = (ZERO_SUM_POS_t >= ZERO_hx)
    If ZERO_D_POS_check Then
        ESITO_ZERO = "'DERIVA POSITIVA'"
        ZERO_D_POS = 0.7 * (ZERO_kx + ZERO_SUM_POS_t / ZERO_N_POS)
    Else
        ZERO_D_POS = 0
    End If
    
    ZERO_D_NEG_check = (ZERO_SUM_NEG_t >= ZERO_hx)
    If ZERO_D_NEG_check Then
        ESITO_ZERO = "'DERIVA NEGATIVA'"
        ZERO_D_NEG = 0.7 * (ZERO_kx + ZERO_SUM_NEG_t / ZERO_N_NEG)
    Else
        ZERO_D_NEG = 0
    End If
    
    '******** verifica di SPAN *******
    SPAN_dt = SpanResult - SpanRif
    SPAN_SUM_POS_p = OLD_SPAN_SUM_POS + SPAN_dt - SPAN_kx
    SPAN_SUM_NEG_p = OLD_SPAN_SUM_NEG - SPAN_dt - SPAN_kx
    
    If SPAN_SUM_POS_p > 0 Then
        SPAN_SUM_POS_t = SPAN_SUM_POS_p
        SPAN_N_POS = OLD_SPAN_N_POS + 1
    Else
        SPAN_SUM_POS_t = 0
        SPAN_N_POS = 0
    End If
    
    If SPAN_SUM_NEG_p > 0 Then
        SPAN_SUM_NEG_t = SPAN_SUM_NEG_p
        SPAN_N_NEG = OLD_SPAN_N_NEG + 1
    Else
        SPAN_SUM_NEG_t = 0
        SPAN_N_NEG = 0
    End If
    
    ESITO_SPAN = "'NESSUNA DERIVA'"
    SPAN_D_POS_check = (SPAN_SUM_POS_t >= SPAN_hx)
    If SPAN_D_POS_check Then
        ESITO_SPAN = "'DERIVA POSITIVA'"
        SPAN_D_POS = 0.7 * (SPAN_kx + SPAN_SUM_POS_t / SPAN_N_POS)
    Else
        SPAN_D_POS = 0
    End If
    
    SPAN_D_NEG_check = (SPAN_SUM_NEG_t >= SPAN_hx)
    If SPAN_D_NEG_check Then
        ESITO_SPAN = "'DERIVA NEGATIVA'"
        SPAN_D_NEG = 0.7 * (SPAN_kx + SPAN_SUM_NEG_t / SPAN_N_NEG)
    Else
        SPAN_D_NEG = 0
    End If
    
    '**************************************************************************
    '*                       fine verifica di DERIVA                          *
    '**************************************************************************
        
    ZeroErr = ZeroResult - ZeroRif
    SpanErr = SpanResult - SpanRif
        
    OraTaratura = CStr(Right("00" & CStr(hour(Now)), 2) & "." & Right("00" & CStr(minute(Now)), 2))
        
    '***** salvataggio dati su mySQL
    With rsQAL3
       strSQL = " INSERT INTO WDS_CALIBRATION ("
       strSQL = strSQL + "cl_system, "
       strSQL = strSQL + "cl_station, "
       strSQL = strSQL + "cl_Hour, "
       strSQL = strSQL + "cl_Date, "
       strSQL = strSQL + "cl_parameter, "
       strSQL = strSQL + "cl_description, "
       strSQL = strSQL + "cl_zero, "
       strSQL = strSQL + "cl_span1, "
       strSQL = strSQL + "cl_tzero, "
       strSQL = strSQL + "cl_tspan1, "
       strSQL = strSQL + "cl_tspan2, "
       strSQL = strSQL + "cl_tspan3, "
       strSQL = strSQL + "cl_tspan4, "
       strSQL = strSQL + "cl_tspan5, "
       strSQL = strSQL + "cl_error, "
       'nik 16/04/2013
       strSQL = strSQL + "cl_tipo, "
       
       For ndx = 1 To 43
            If ndx < 43 Then
                strSQL = strSQL + "c" & ndx & ", "
            Else
                strSQL = strSQL + "c" & ndx
            End If
       Next
        
        strSQL = strSQL + ") "
        strSQL = strSQL + "VALUES ("
        strSQL = strSQL + .ParSQLStr(CStr(gsImpianto)) & ","
        strSQL = strSQL + .ParSQLStr(CStr(gsClienteDi)) & ","
        strSQL = strSQL + .ParSQLStr(CStr(OraTaratura)) & ","
        strSQL = strSQL + .ParSQLDate(CDate(Now)) & ","
        strSQL = strSQL + .ParSQLStr(ParametriStrumenti(vnIndiceParametro).NomeParametro) & ","
        strSQL = strSQL + .ParSQLStr(ParametriStrumenti(vnIndiceParametro).DescrParametro & " " & ParametriStrumenti(vnIndiceParametro).UnitaMisura) & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(ZeroResult, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(SpanResult, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(ZeroErr, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(SpanErr, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(ZeroRif, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(SpanRif, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(ZeroResult, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + Replace(CStr(FormatNumber(SpanResult, 2, -2, -2, 0)), ",", ".") & ","
        strSQL = strSQL + .ParSQLStr(CStr(ErroreStr)) & ","
        'nik 16/04/2013
        strSQL = strSQL + .ParSQLStr(CStr(TipoCal)) & ","
              
       For ndx = 1 To 43
          Select Case ndx
            '***** precisione ******
            Case 1
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_hs, 2, -2, -2, 0)), ",", ".") & ","
            Case 2
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_hs, 2, -2, -2, 0)), ",", ".") & ","
            Case 3
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_ks, 2, -2, -2, 0)), ",", ".") & ","
            Case 4
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_ks, 2, -2, -2, 0)), ",", ".") & ","
            Case 5
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_dt, 2, -2, -2, 0)), ",", ".") & ","
            Case 6
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_dt, 2, -2, -2, 0)), ",", ".") & ","
            Case 7
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_sp, 2, -2, -2, 0)), ",", ".") & ","
            Case 8
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_sp, 2, -2, -2, 0)), ",", ".") & ","
            Case 9
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_st, 2, -2, -2, 0)), ",", ".") & ","
            Case 10
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_st, 2, -2, -2, 0)), ",", ".") & ","
            Case 11
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_Nst, 0, -2, -2, 0)), ",", ".") & ","
            Case 12
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_Nst, 0, -2, -2, 0)), ",", ".") & ","
            Case 13
                If ZERO_check Then
                    strSQL = strSQL & "'SI'" & ","
                Else
                    strSQL = strSQL & "'NO'" & ","
                End If
            Case 14
                If SPAN_check Then
                    strSQL = strSQL & "'SI'" & ","
                Else
                    strSQL = strSQL & "'NO'" & ","
                End If
            '******* deriva *******
            Case 20
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_hx, 2, -2, -2, 0)), ",", ".") & ","
            Case 21
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_hx, 2, -2, -2, 0)), ",", ".") & ","
            Case 22
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_kx, 2, -2, -2, 0)), ",", ".") & ","
            Case 23
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_kx, 2, -2, -2, 0)), ",", ".") & ","
            Case 24
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_dt, 2, -2, -2, 0)), ",", ".") & ","
            Case 25
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_dt, 2, -2, -2, 0)), ",", ".") & ","
            Case 26
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_SUM_POS_p, 2, -2, -2, 0)), ",", ".") & ","
            Case 27
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_SUM_POS_t, 2, -2, -2, 0)), ",", ".") & ","
            Case 28
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_N_POS, 0, -2, -2, 0)), ",", ".") & ","
            Case 29
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_SUM_NEG_p, 2, -2, -2, 0)), ",", ".") & ","
            Case 30
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_SUM_NEG_t, 2, -2, -2, 0)), ",", ".") & ","
            Case 31
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_N_NEG, 0, -2, -2, 0)), ",", ".") & ","
            Case 32
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_SUM_POS_p, 2, -2, -2, 0)), ",", ".") & ","
            Case 33
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_SUM_POS_t, 2, -2, -2, 0)), ",", ".") & ","
            Case 34
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_N_POS, 0, -2, -2, 0)), ",", ".") & ","
            Case 35
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_SUM_NEG_p, 2, -2, -2, 0)), ",", ".") & ","
            Case 36
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_SUM_NEG_t, 2, -2, -2, 0)), ",", ".") & ","
            Case 37
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_N_NEG, 0, -2, -2, 0)), ",", ".") & ","
            Case 38
                strSQL = strSQL + CStr(ESITO_ZERO) & ","
            Case 39
                strSQL = strSQL + CStr(ESITO_SPAN) & ","
            Case 40
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_D_POS, 2, -2, -2, 0)), ",", ".") & ","
            Case 41
                strSQL = strSQL + Replace(CStr(FormatNumber(ZERO_D_NEG, 2, -2, -2, 0)), ",", ".") & ","
            Case 42
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_D_POS, 2, -2, -2, 0)), ",", ".") & ","
            Case 43
                strSQL = strSQL + Replace(CStr(FormatNumber(SPAN_D_NEG, 2, -2, -2, 0)), ",", ".")
            '*************************
            Case Else
                If ndx < 43 Then
                    strSQL = strSQL & "0, "
                Else
                    strSQL = strSQL & "0"
                End If
          End Select
        Next
        strSQL = strSQL + ")"
        
        'luca 08/11/2016 salvo QAL3 solo se non è il client
        If Not Client Then
            .ExecuteSql (strSQL)
        End If
    End With
    Set rsQAL3 = Nothing
    
    Dim strInizioTag As String
    strInizioTag = CStr(NumeroLinea) + ".AM" & Format(ParametriStrumenti(vnIndiceParametro).CodiceParametro, "000")
    'Nicolò Luglio 2016
    Call ScriviTag(strInizioTag & "_QAL3_DATE", CStr(DateSerial(year(Now), month(Now), day(Now))))
    Call ScriviTag(strInizioTag & "_QAL3_HOUR", Format(Now, "hh.nn.ss"))
    
    'luca 15/09/2016 aggiungo scrittura dei risultati di zero e span
    Call ScriviTag(strInizioTag & "_QAL3_ZERORES", ZeroResult)
    Call ScriviTag(strInizioTag & "_QAL3_SPANRES", SpanResult)
    
    'Nicolò Luglio 2016
    If InStr(ESITO_ZERO, "NESSUNA") > 0 And InStr(ESITO_SPAN, "NESSUNA") > 0 Then
        Call ScriviTag(strInizioTag + "_QAL3_TOTRESULT", "OK")
    Else
        Call ScriviTag(strInizioTag + "_QAL3_TOTRESULT", "Errore")
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("TaratureSalvaQAL3 " + Error(Err), 1, OPC)
    Resume Next

End Sub

'luca 04/10/2016
Private Sub GestioneResetQAL3(Indice As Integer)

    On Error GoTo Gesterrore

    If LeggiTag(CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(Indice).CodiceParametro, "000") & "_QAL3_RESET") = True Then
        Call AzzeraStatisticaQAL3(ParametriStrumenti(Indice).NomeParametro)
        Call ScriviTag(CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(Indice).CodiceParametro, "000") + "_QAL3_RESET", 0)
    End If

    Exit Sub
    
Gesterrore:
    Call WindasLog("GestioneResetQAL3 " + Error(Err), 1, OPC)
End Sub

Private Sub AzzeraStatisticaQAL3(ByVal NomePar)

    Dim rsQAL3 As Object
    Dim Data, Ora
    
    On Error GoTo Gesterrore
    
    'Alby Febbraio 2016
    Call WindasLog("Reset QAL3 per parametro: " + NomePar, 0, OPC)
    
    '***** Lettura parametri di configurazione database su file bfdesk.xml *****
    Call GetConnectionParam
                
    NewDataObj rsQAL3
    With rsQAL3
        '**** estrae l'ultimo record calibrazione per lo strumento/stazione ****
        strSQL = "SELECT * FROM WDS_CALIBRATION WHERE CL_TIPO='Q' AND CL_PARAMETER = '" & CStr(NomePar) & "' AND CL_STATION = '" & CStr(NumeroLinea) & "_SiCEMS' ORDER BY cl_date DESC, cl_hour DESC"
        .selectionfast CStr(strSQL)
        If Not .iseof Then
          '*************** estrae data e ora del record da azzerare **************
          Data = .getValue("cl_Date")
          Ora = .getValue("cl_hour")
                  
          '********************* azzera i parametriQAL3 del record ***************
          '********** PRECISIONE *******
          'luca 13/10/2016
          strSQL = "UPDATE WDS_CALIBRATION SET  c5=0, c6=0, c9=0, c10=0, c11=0, c12=0, c13='Reset QAL3', c14='Reset QAL3' WHERE CL_STATION = '" & CStr(NumeroLinea) & "_SiCEMS' AND CL_PARAMETER = '" & CStr(NomePar) & "' AND cl_date='" & Data & "' AND cl_hour='" & Ora & "'"
          .ExecuteSql (CStr(strSQL))
          
          '********* DERIVA ************
            'luca 13/10/2016
          strSQL = "UPDATE WDS_CALIBRATION SET c24=0, c25=0, c27=0, c28=0, c30=0, c31=0, c33=0, c34=0, c36=0, c37=0, c38='Reset QAL3', c39='Reset QAL3', c40=0, c41=0, c42=0, c43=0 WHERE CL_STATION = '" & CStr(NumeroLinea) & "_SiCEMS' AND CL_PARAMETER = '" & NomePar & "' AND cl_date='" & Data & "' AND cl_hour='" & Ora & "'"
          .ExecuteSql (CStr(strSQL))
        End If
    End With
    Set rsQAL3 = Nothing
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("AzzeraStatisticaQAL3 " + Error(Err), 1, OPC)

End Sub


