Attribute VB_Name = "Personalizzazioni"
Option Explicit

Public Type ENELMisure
    Sequenza As Integer
    indice As Integer
    AI_Zero As String
    AI_Span As String
    Selettore_Zero As String
    Selettore_Span As String
End Type
Dim Misure(4) As ENELMisure

'Federica luglio 2018
Sub ControlloWatchdogPLCMater()
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
            valore = LeggiTag(CStr(NumeroLinea) & " DI30")
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
        manValoreDigitale(999, 1) = 1
        RecuperoFatto = False
    Else
        manValoreDigitale(999, 1) = 0
        
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
    Call WindasLog("ControlloWatchdogPLCMater: " + Error(Err), 1, OPC)
 
End Sub

'luca maggio 2018
Sub AttivaUsciteSuperoMetaEnergia()

    Dim ElencoAllarmi() As String
    Dim i As Integer
    
    On Error GoTo Gesterrore
    ElencoAllarmi = Split(Trim(Generiche(iElencoAllarmiSuperoPLC).Testo), ";")
    
    For i = 0 To UBound(ElencoAllarmi)
        If Valore_DI(CInt(ElencoAllarmi(i))) = 1 Then
            ScriviTag IIf(NumeroLinea = 1, "DO16", "DO17"), 1
            Exit Sub
        End If
    Next i
    
    ScriviTag IIf(NumeroLinea = 1, "DO16", "DO17"), 0
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("StatoImpianto " + Error(Err), 1, OPC)

End Sub

'luca maggio 2018
Public Sub CalcolaPortataPitotMetaEnergia()
    
    Dim ValoreDeltaP As Double
    Dim ValoreNumeratoreSottoRadice As Double
    Dim ValoreDenominatoreSottoRadice As Double
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri ****
    If IngressoPortata < 0 Or IngressoPress < 0 Or IngressoTemp < 0 Then
        Call WindasLog("CalcolaPortataPitotKSenzaArea: Ingressi non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    '**** Calcolo ****
    If ParametriStrumenti(IngressoPortata).Acquisizione Then
        ValoreDeltaP = ValIst(0, IngressoPortata)
        
        ValoreNumeratoreSottoRadice = 2 * 100 * ValoreDeltaP * ValIst(0, IngressoTemp) * 273.15
        ValoreDenominatoreSottoRadice = 1.29 * (273.15 + ValIst(0, IngressoTemp)) * 1013.25
        
        If ValoreNumeratoreSottoRadice >= 0 And ValoreDenominatoreSottoRadice > 0 Then
            ValIst(0, IngressoPortata) = 3600 * Val(Replace(Trim(Generiche(Kportata).Par), ",", ".")) * Val(Replace(Trim(Generiche(AreaCamino).Par), ",", ".")) * Sqr(ValoreNumeratoreSottoRadice / ValoreDenominatoreSottoRadice)
            ValIst(1, IngressoPortata) = ValIst(0, IngressoPortata)
        Else
            ValIst(0, IngressoPortata) = -9999
            ValIst(1, IngressoPortata) = -9999
            Status(0, IngressoPortata) = "ERR"
            Status(1, IngressoPortata) = "ERR"
        End If
    End If
    
    Exit Sub
Gesterrore:
    Call WindasLog("CalcolaPortataPitotMetaEnergia: " & Error(Err()), 1, "OPC")
End Sub

Public Sub MaterBioStatoImpianto()

    On Error GoTo Gesterrore
    
    Dim CodiceStatoImpianto As Integer
    Dim iIndice As Integer
    
    On Error GoTo Gesterrore
        
    If IngressoStatoImpianto = -1 Then Exit Sub
    
    CodiceStatoImpianto = 34
    'Verifico se sono il marcia
    If Valore_DI(16) = 1 Then
        CodiceStatoImpianto = 30
        If Valore_DI(17) = 1 Then CodiceStatoImpianto = 36
        If Valore_DI(18) = 1 Then CodiceStatoImpianto = 32
        If Valore_DI(19) = 1 Then CodiceStatoImpianto = 31
    End If
        
    If Trim(Generiche(iMisureSimulate).Testo) <> "" Then
        'Se ho misure simulate forzo anche lo stato impianto
        ValIst(0, IngressoStatoImpianto) = 30
    Else
        ValIst(0, IngressoStatoImpianto) = CodiceStatoImpianto
    End If
    Status(0, IngressoStatoImpianto) = "VAL"
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("MaterBioStatoImpianto: " & Error(Err()), 1, "OPC")

End Sub

Sub ENELControlloTarature()
'TODO: Nella parametrizzazione inserire stringa con le misure coinvolte (es. 0;2;8)

    Dim i As Integer
    
    On Error GoTo Gesterrore
    
    '**** Carico le misure ****
    With Misure(0)  'CO
        .Sequenza = 1
        .indice = 0
        .AI_Zero = "11"
        .AI_Span = "12"
        .Selettore_Zero = "16"
        .Selettore_Span = "24"
    End With
    With Misure(1)  'NOX
        .Sequenza = 1
        .indice = 11
        .AI_Zero = "13"
        .AI_Span = "14"
        .Selettore_Zero = "17"
        .Selettore_Span = "25"
    End With
    With Misure(2)  'O2
        .Sequenza = 1
        .indice = 2
        .AI_Zero = "15"
        .AI_Span = "16"
        .Selettore_Zero = "18"
        .Selettore_Span = "26"
    End With
    With Misure(3)  'SO2
        .Sequenza = 1
        .indice = 3
        .AI_Zero = "17"
        .AI_Span = "18"
        .Selettore_Zero = "19"
        .Selettore_Span = "27"
    End With
    With Misure(4)  'THC
        .Sequenza = 2
        .indice = 9
        .AI_Zero = "19"
        .AI_Span = "20"
        .Selettore_Zero = ""
        .Selettore_Span = ""
    End With
'    For i = 0 To UBound(Misure)
'        Call GestioneResetQAL3(CInt(Misure(i)))
'    Next i
        
    '                 Sequenza, QAL3incorso, QAL3finitaOK
    Call ENELLeggiVariabiliWinCCperTarature(1, 86, 87)
    Call ENELLeggiVariabiliWinCCperTarature(2, 90, 91)

    Exit Sub
    
Gesterrore:
    Call WindasLog("ENELControlloTarature " + Error(Err), 1, OPC)

End Sub

Private Sub ENELLeggiVariabiliWinCCperTarature(Sequenza, QAL3inCorso, QAL3ultimata)
'TODO: Eventualmente configurare le Tag se la prevedono

    Static statoQAL3(2) As Integer
    Static cont(2) As Integer
    'luca aprile 2017
    Dim i As Integer
    Dim tempQAL3(9) As Double
    Dim tempQAL3Selettori(9) As Boolean
    
    On Error GoTo Gesterrore
    
    If Valore_DI(QAL3inCorso) = 1 Then
        statoQAL3(Sequenza) = 1
    End If
    
    If statoQAL3(Sequenza) = 1 Then
        If Valore_DI(QAL3ultimata) = 1 Then
            'luca 22/09/2016 inserisco contatore perchè salva i risultati troppo presto (nuovi risultati non ancora a disposizione lato PLC)
            If cont(Sequenza) = 0 Then
            
                Call WindasLog("QAL3 terminata regolarmente...  salvataggio risultati", 0, OPC)
                
                For i = 0 To UBound(Misure)
                    Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(0), span(0), CStr(NumeroLinea) & "AI" & Misure(i).AI_Zero, CStr(NumeroLinea) & "AI" & Misure(i).AI_Span)
                    If (Misure(i).Selettore_Zero <> "") And (Misure(i).Selettore_Span <> "") Then
                        ParamSelected(Misure(i).indice) = CBool(LeggiTag(CStr(NumeroLinea) & "DO" & Misure(i).Selettore_Zero)) And CBool(LeggiTag(CStr(NumeroLinea) & "DO" & Misure(i).Selettore_Span))  'Nicolò Agosto 2016
                    Else
                        ParamSelected(Misure(i).indice) = True
                    End If
                Next i
                
'                'CO
'                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(0), span(0), IIf(NumeroLinea = 1, "AI20", "AI26"), IIf(NumeroLinea = 1, "AI21", "AI27"))
'
'                ParamSelected(0) = CBool(LeggiTag("DI136")) And CBool(LeggiTag("DI144"))  'Nicolò Agosto 2016
'
'                Call WindasLog("Parametro CO selettori: " & ParamSelected(0), 0, OPC)
'
'                'NOx
'                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(1), span(1), IIf(NumeroLinea = 1, "AI22", "AI28"), IIf(NumeroLinea = 1, "AI23", "AI29"))
'
'                ParamSelected(1) = CBool(LeggiTag("DI137")) And CBool(LeggiTag("DI145"))   'Nicolò Agosto 2016
'                'ParamSelected(0) = True
'                Call WindasLog("Parametro NOX selettori: " & ParamSelected(1), 0, OPC)
'
'                'O2
'                Call LeggiVariabiliWinCCperTaratureCaricamentoRisultatiPLC(zero(2), span(2), IIf(NumeroLinea = 1, "AI24", "AI30"), IIf(NumeroLinea = 1, "AI25", "AI31"))
'
'                ParamSelected(2) = CBool(LeggiTag("DI138")) And CBool(LeggiTag("DI146"))   'Nicolò Agosto 2016
'                'ParamSelected(2) = True
'                Call WindasLog("Parametro O2 selettori: " & ParamSelected(2), 0, OPC)
                        
                Call TaratureSalvaQAL3(Misure(i).indice, "Q")
'                Call TaratureSalvaQAL3(1, "Q")
'                Call TaratureSalvaQAL3(2, "Q")
                        
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
    Call WindasLog("ENELLeggiVariabiliWinCCperTarature Sequenza: " + Format(Sequenza, "0") + " " + Error(Err), 1, OPC)

End Sub

Sub EnelStatoImpianto()
'§ Determinazione stato impianto da digitale

    Dim CodiceStatoImpianto As Integer
    Dim iIndice As Integer
    
    On Error GoTo Gesterrore
        
    If IngressoStatoImpianto = -1 Then Exit Sub
    
    CodiceStatoImpianto = 34
    'Verifico se sono in marcia
    If Valore_DI(102) = 1 Then
        CodiceStatoImpianto = 30
        'Verifico se sono sopra il minimo tecnico
        If Valore_DI(103) <> 0 Then CodiceStatoImpianto = 31
    End If
    
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

'Federica novembre 2018 - Calcolo QFumi da Delta_p
Public Sub ENELCalcolaPortataDaDeltaP()

    Dim ValorePortata As Double
    Dim ValoreDeltaP As Double
    Dim FlagDeltaP As String
    Dim Coefficiente_K As Double
    Dim ValorePressione As Double
    Dim FlagPressione As String
    Dim ValoreTemperatura As Double
    Dim FlagTemperatura As String
    Dim FattoreEspansione As Double
    Dim Densita As Double
    Dim AreaCamino As Double

    On Error GoTo Gesterrore
    
    '**** Verifica parametri necessari ****
    If IngressoPortata <= 0 Then
        Call WindasLog("CalcolaPortataDaDeltaP: IngressoPortata non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngressoDeltaP <= 0) Or (TrasformaInDbl(Generiche(Kportata).Par) <= 0) Or (IngressoPress <= 0) Or (IngressoTemp <= 0) Then
        ValIst(0, IngressoPortata) = -9999
        Status(0, IngressoPortata) = "ERR"
        ValIst(1, IngressoPortata) = -9999
        Status(1, IngressoPortata) = "ERR"
        
        Call WindasLog("ENELCalcolaPortataDaDeltaP: IngressoDeltaP, IngrassoPress, IngressoTemp, K non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    If ParametriStrumenti(IngressoPortata).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(IngressoPortata).Acquisizione Then
            ValoreDeltaP = ValIst(0, IngressoDeltaP)
            FlagDeltaP = Status(0, IngressoDeltaP)
            Coefficiente_K = TrasformaInDbl(Generiche(Kportata).Par)
            ValorePressione = ValIst(0, IngressoPress)
            FlagPressione = Status(0, IngressoPress)
            ValoreTemperatura = ValIst(0, IngressoTemp)
            FlagTemperatura = Status(0, IngressoTemp)
            
            If (ValoreDeltaP <> -9999) And (Valido(FlagDeltaP)) And (Coefficiente_K > 0) And (ValorePressione <> -9999) And (Valido(FlagPressione)) And (ValoreTemperatura <> -9999) And (Valido(FlagTemperatura)) Then
                
                ValorePortata = Coefficiente_K * ((Sqr(ValorePressione)) / (Sqr(273.15 + ValoreTemperatura))) * Sqr(ValoreDeltaP)
                
                ValIst(0, IngressoPortata) = ValorePortata
                ValIst(1, IngressoPortata) = ValIst(0, IngressoPortata)
                Status(0, IngressoPortata) = "VAL"
                Status(1, IngressoPortata) = "VAL"
            Else
                ValIst(0, IngressoPortata) = 0
                ValIst(1, IngressoPortata) = ValIst(0, IngressoPortata)
                Status(0, IngressoPortata) = "ERR"
                Status(1, IngressoPortata) = "ERR"
            End If
        End If
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("CalcolaPortataDaDeltaP: " & Error(Err()), 1, "OPC")

End Sub

