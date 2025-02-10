Attribute VB_Name = "SMECLoud"
Option Explicit

'Edoardo 21/11/2016
Private Sub SettaAllarmeSMECloud(NumeroAllarmeSMECloud As Integer, IndiceParametro As Integer, Condizione As Boolean)
    
    Dim i As Integer
    
    On Error GoTo GestErrore

    i = ParametriStrumenti(IndiceParametro).idDatabase
    If Condizione Then manValoreDigitale(i, NumeroAllarmeSMECloud) = 1 Else manValoreDigitale(i, NumeroAllarmeSMECloud) = 0
    
    Exit Sub
    
GestErrore:
    Call WindasLog("SettaAllarmeSMECloud: " + Error(Err), 1, OPC)
End Sub

'Federica dicembre 2017 - Gestione DO per allarmi supero limiti medie
Public Sub GestioneSMECloudDOAllarmiMedie()

    Dim Parametri() As String
    Dim Allarmi() As String
    Dim Ret(99)
    Dim iParametro As Integer
    Dim iAllarme As Integer
    Dim iRet As Integer
    Dim iDO As Integer
    'Federica dicembre 2017
    Dim SMECloudElencoParametri As String
    Dim SMECloudElencoAllarmi As String

    On Error GoTo GestErrore
    
    SMECloudElencoParametri = Trim(Generiche(iSMECloudParametri).Testo)
    SMECloudElencoAllarmi = Trim(Generiche(iSMECloudAllarmi).Testo)
    
    If SMECloudElencoParametri = "" Then Exit Sub   'Non sono stati configurati parametri da controllare
    If SMECloudElencoAllarmi = "" Then Exit Sub   'Non sono stati configurati allarmi da controllare
    
    '***** Configurazione *****
    Parametri = Split(SMECloudElencoParametri, ";")
    Allarmi = Split(SMECloudElencoAllarmi, ";")
    
    '***** Scansione *****
    For iAllarme = 0 To UBound(Allarmi)
        For iParametro = 0 To UBound(Parametri)
            If IndiceDO(CInt(Parametri(iParametro)), CInt(Allarmi(iAllarme))) <> "" Then    'Se ho l'indice della DO configurato
                iDO = IndiceDO(CInt(Parametri(iParametro)), CInt(Allarmi(iAllarme)))
                If Ret(iDO) = vbEmpty Then Ret(iDO) = 0 'Inizializzo il valore della DO
                If manValoreDigitale(CInt(Parametri(iParametro)), CInt(Allarmi(iAllarme))) = 1 Then
                    Ret(iDO) = 1
                    Exit For
                End If
            End If
        Next iParametro
    Next iAllarme
    
    '***** Scrittura ******
    For iRet = 0 To UBound(Ret)
        If Ret(iRet) <> "" Then ScriviTag CStr(NumeroLinea) & "_DB80_DO_" & Format(iRet, "00"), Ret(iRet)
    Next iRet
    
    Exit Sub
GestErrore:
    Call WindasLog("SetVariabiliWinCCAllarmiMedie: " & Error(Err()), 1, "OPC")

End Sub

Sub GestioneSMECloud()

    Dim iIDParametro As Integer
    Dim IndiceMorsettoDig As Long
    Dim Condizione As Boolean
    Dim ValStatoImpianto
    
    Dim SMELimiteOraSemi As Double
    Dim SMETagOraSemi As String
    
    On Error GoTo GestErrore

    'Gestisco le differenze con ORE_SEMIORE

    SMELimiteOraSemi = IIf(OreSemiore = 1, ParametriStrumenti(iIDParametro).LimConcMediaOraria, ParametriStrumenti(iIDParametro).LimConcMediaSemiorariaA)
    ValStatoImpianto = LeggiTag(CStr(NumeroLinea) & ".AM" & Format(ParametriStrumenti(IngressoStatoImpianto).CodiceParametro, "000") & IIf(OreSemiore = 1, "_MONU", "_MSNU"))

    '***** Ciclo di Tutti i Possibili Parametri degli Strumenti *****
    For iIDParametro = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(iIDParametro)
            If SMELimiteOraSemi > 0 Then
                If Val(Replace(MediaOraInCorso(1, iIDParametro), ",", ".")) <> -9999 Then
                    'Allarmi per supero soglia attenzione media oraria/semioraria in costruzione (100)
                    If .SogliaAttenzione > 0 Then
                        Condizione = (Val(Replace(MediaOraInCorso(1, iIDParametro), ",", ".")) > .SogliaAttenzione)
                        Call SettaAllarmeSMECloud(100, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(100, iIDParametro, False)
                    End If
                    
                    'Allarmi per supero soglia allarme media oraria/semioraria in costruzione (101)
                    If .SogliaAllarme > 0 Then
                        Condizione = (Val(Replace(MediaOraInCorso(1, iIDParametro), ",", ".")) > .SogliaAllarme)
                        Call SettaAllarmeSMECloud(101, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(101, iIDParametro, False)
                    End If
                    
                    'Allarmi per supero limite media oraria/semioraria in costruzione (102)
                    Condizione = (Val(Replace(MediaOraInCorso(1, iIDParametro), ",", ".")) > SMELimiteOraSemi)
                    Call SettaAllarmeSMECloud(102, iIDParametro, Condizione)
                Else
                    Call SettaAllarmeSMECloud(100, iIDParametro, False)
                    Call SettaAllarmeSMECloud(101, iIDParametro, False)
                    Call SettaAllarmeSMECloud(102, iIDParametro, False)
                End If
                
                'Allarmi per supero limite ultima media oraria (112)
                '(LeggiTag(CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & "_MONU_VAL") = "VAL" Or LeggiTag(CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & "_MONU_VAL") = "AUX")
                SMETagOraSemi = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & IIf(OreSemiore = 1, "_MONU", "_MSNU")
                If Valido(LeggiTag(SMETagOraSemi & "_VAL")) And Val(Replace(ValStatoImpianto, ",", ".")) >= 70 Then
                    Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > SMELimiteOraSemi)
                    Call SettaAllarmeSMECloud(112, iIDParametro, Condizione)
                Else
                     Call SettaAllarmeSMECloud(112, iIDParametro, False)
                End If
                
                'Allarmi per invalidità ultima media oraria (113)
                'LeggiTag(CStr(NumeroLinea) & ".AM" & Format(CodiceParametro(iIDParametro), "000") & "_MONU_VAL") <> "VAL" And LeggiTag(CStr(NumeroLinea) & ".AM" & Format(CodiceParametro(iIDParametro), "000") & "_MONU_VAL") <> "AUX")
                Condizione = (Val(Replace(ValStatoImpianto, ",", ".")) >= 70 And Valido(LeggiTag(SMETagOraSemi & "_VAL")))
                Call SettaAllarmeSMECloud(113, iIDParametro, Condizione)
            Else
                Call SettaAllarmeSMECloud(100, iIDParametro, False)
                Call SettaAllarmeSMECloud(101, iIDParametro, False)
                Call SettaAllarmeSMECloud(102, iIDParametro, False)
                Call SettaAllarmeSMECloud(112, iIDParametro, False)
                Call SettaAllarmeSMECloud(113, iIDParametro, False)
            End If
            
            If .LimConcMediaGiornaliera > 0 Then
                SMETagOraSemi = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & IIf(OreSemiore = 1, "_MGONC", "_MGSNC")
                If Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) <> -9999 Then
                    'Allarmi per supero soglia attenzione media giornaliera in costruzione (120)
                    If .SogliaAttenzioneGiornaliera > 0 Then
                        Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .SogliaAttenzioneGiornaliera)
                        Call SettaAllarmeSMECloud(120, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(120, iIDParametro, False)
                    End If
                    
                    'Allarmi per supero soglia allarme media giornaliera in costruzione (121)
                    If .SogliaAllarmeGiornaliera > 0 Then
                        Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .SogliaAllarmeGiornaliera)
                        Call SettaAllarmeSMECloud(121, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(121, iIDParametro, False)
                    End If
                Else
                    Call SettaAllarmeSMECloud(120, iIDParametro, False)
                    Call SettaAllarmeSMECloud(121, iIDParametro, False)
                End If
                
                'Federica gennaio 2018
                'Allarmi per supero limite media giornaliera in corso (131)
                '(LeggiTag(CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & "_MGONC_VAL") = "VAL" Or LeggiTag(CStr(NumeroLinea) & ".AM" & Format(CodiceParametro(iIDParametro), "000") & "_MGONC_VAL") = "AUX")
                If Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) <> -9999 And Valido(LeggiTag(SMETagOraSemi & "_VAL")) Then
                    Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .LimConcMediaGiornaliera)
                    Call SettaAllarmeSMECloud(131, iIDParametro, Condizione)
                Else
                    Call SettaAllarmeSMECloud(131, iIDParametro, False)
                End If
                
                'Allarmi per supero limite ultima media giornaliera (132)
                '(LeggiTag(CStr(NumeroLinea) & ".AM" & Format(CodiceParametro(iIDParametro), "000") & "_MGONU_VAL") = "VAL" Or LeggiTag(CStr(NumeroLinea) & ".AM" & Format(CodiceParametro(iIDParametro), "000") & "_MGONU_VAL") = "AUX")
                SMETagOraSemi = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & IIf(OreSemiore = 1, "_MGONU", "_MGSNU")
                If Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) <> -9999 And Valido(LeggiTag(SMETagOraSemi & "_VAL")) Then
                    Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .LimConcMediaGiornaliera)
                    Call SettaAllarmeSMECloud(132, iIDParametro, Condizione)
                Else
                    Call SettaAllarmeSMECloud(132, iIDParametro, False)
                End If
            Else
                Call SettaAllarmeSMECloud(120, iIDParametro, False)
                Call SettaAllarmeSMECloud(121, iIDParametro, False)
                Call SettaAllarmeSMECloud(132, iIDParametro, False)
            End If
            
            'Federica ottobre 2017 - Allarmi per soglie mensili
            '**** SOGLIE MENSILI ****
            If .LimConcMediaMensile > 0 Then
                SMETagOraSemi = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & IIf(OreSemiore = 1, "_MMONC", "_MMSNC")
                If Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) <> -9999 Then
                    'Allarmi per supero soglia attenzione media mensile in costruzione (140)
                    If .SogliaAttenzioneMensile > 0 Then
                        Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .SogliaAttenzioneMensile)
                        Call SettaAllarmeSMECloud(140, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(140, iIDParametro, False)
                    End If
                    
                    'Allarmi per supero soglia allarme media mensile in costruzione (141)
                    If .SogliaAllarmeMensile > 0 Then
                        Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .SogliaAllarmeMensile)
                        Call SettaAllarmeSMECloud(141, iIDParametro, Condizione)
                    Else
                        Call SettaAllarmeSMECloud(141, iIDParametro, False)
                    End If
                Else
                    Call SettaAllarmeSMECloud(140, iIDParametro, False)
                    Call SettaAllarmeSMECloud(141, iIDParametro, False)
                End If
                
                'Allarmi per supero limite ultima media mensile (152)
                SMETagOraSemi = CStr(NumeroLinea) & ".AM" & Format(.CodiceParametro, "000") & IIf(OreSemiore = 1, "_MMONU", "_MMSNU")
                If Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) <> -9999 And Valido(LeggiTag(SMETagOraSemi & "_VAL")) Then
                    Condizione = (Val(Replace(LeggiTag(SMETagOraSemi), ",", ".")) > .LimConcMediaMensile)
                    Call SettaAllarmeSMECloud(152, iIDParametro, Condizione)
                Else
                    Call SettaAllarmeSMECloud(152, iIDParametro, False)
                End If
            Else
                Call SettaAllarmeSMECloud(140, iIDParametro, False)
                Call SettaAllarmeSMECloud(141, iIDParametro, False)
                Call SettaAllarmeSMECloud(152, iIDParametro, False)
            End If
        End With
    Next iIDParametro
    
    Exit Sub
    
GestErrore:
    Call WindasLog("GestioneSMECloud: " + Error(Err), 1, OPC)
End Sub

