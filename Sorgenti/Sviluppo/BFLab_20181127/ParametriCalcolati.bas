Attribute VB_Name = "ParametriCalcolati"
Option Explicit

'Federica settembre 2017 - Calcolo H2O
Public Sub CalcolaH2O(Optional ByVal FormulaAlternativa As Boolean = False)

    Dim ValoreO2Umido As Double
    Dim ValoreO2 As Double
    Dim FlagO2Umido As String
    Dim FlagO2 As String
    Dim ValoreH2O As Double
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri necessari ****
    If IngressoH2O <= 0 Then
        Call WindasLog("CalcolaH2O: IngressoH2O non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngressoO2Umido <= 0) Or (IngressoO2 <= 0) Then
        ValIst(0, IngressoH2O) = -9999
        Status(0, IngressoH2O) = "ERR"
        ValIst(1, IngressoH2O) = -9999
        Status(1, IngressoH2O) = "ERR"
        
        Call WindasLog("CalcolaH2O: IngressoO2Umido o IngressoO2 non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    If ParametriStrumenti(IngressoH2O).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(IngressoH2O).Acquisizione Then
            ValoreO2Umido = ValIst(0, IngressoO2Umido)
            ValoreO2 = CalcolaQAL2(IngressoO2, ValIst(0, IngressoO2))
            FlagO2Umido = Status(0, IngressoO2Umido)
            FlagO2 = Status(0, IngressoO2)
            
            If (ValoreO2Umido <> -9999) And (ValoreO2 <> -9999) And (InStr("VAL AUX", FlagO2Umido) > 0) And (InStr("VAL AUX", FlagO2) > 0) Then
                If ValoreO2 <> 0 Then
                    'Federica dicembre 2017 - Aggiunta formula alternativa
                    If FormulaAlternativa Then
                        ValoreH2O = 100 * (ValoreO2 - ValoreO2Umido) / ValoreO2
                    Else
                        ValoreH2O = 100 - (ValoreO2Umido / ValoreO2 * 100)
                    End If
                    If ValoreH2O < 0 Then ValoreH2O = 0
                Else
                    ValoreH2O = 0
                End If
                
                ValIst(0, IngressoH2O) = ValoreH2O
                ValIst(1, IngressoH2O) = ValIst(0, IngressoH2O)
                Status(0, IngressoH2O) = "VAL"
                Status(1, IngressoH2O) = "VAL"
            Else
                ValIst(0, IngressoH2O) = -9999
                ValIst(1, IngressoH2O) = -9999
                Status(0, IngressoH2O) = "ERR"
                Status(1, IngressoH2O) = "ERR"
            End If
        End If
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("CalcolaH2O: " & Error(Err), 1, "OPC")

End Sub

'luca maggio 2018
Public Sub CalcolaNOxNH3Istantaneo(TipoDato As Integer)

    Dim ValoreNOx As Double
    Dim ValoreNH3 As Double
    Dim FlagNOx As String
    Dim FlagNH3 As String
    Dim ValoreNOxNH3 As Double
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri necessari ****
    If IngressoNOXNH3 <= 0 Then
        Call WindasLog("CalcolaNOxNH3Istantaneo: IngressoNOXNH3 non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngressoNOX <= 0) Or (IngressoNH3 <= 0) Then
        ValIst(TipoDato, IngressoNOXNH3) = -9999
        Status(TipoDato, IngressoNOXNH3) = "ERR"
        
        Call WindasLog("CalcolaNOxNH3Istantaneo: IngressoNOX o IngressoNH3 non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    If ParametriStrumenti(IngressoNOXNH3).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(IngressoNOXNH3).Acquisizione Then
            ValoreNOx = ValIst(TipoDato, IngressoNOX)
            ValoreNH3 = ValIst(TipoDato, IngressoNH3)
            FlagNOx = Status(TipoDato, IngressoNOX)
            FlagNH3 = Status(TipoDato, IngressoNH3)
            
            If (ValoreNOx <> -9999) And (ValoreNH3 <> -9999) And (InStr("VAL AUX", FlagNOx) > 0) And (InStr("VAL AUX", FlagNH3) > 0) Then
                ValoreNOxNH3 = ValIst(TipoDato, IngressoNOX) + (ValIst(TipoDato, IngressoNH3) * 2.7)
                
                ValIst(TipoDato, IngressoNOXNH3) = ValoreNOxNH3
                Status(TipoDato, IngressoNOXNH3) = "VAL"
            Else
                ValIst(TipoDato, IngressoNOXNH3) = -9999
                Status(TipoDato, IngressoNOXNH3) = "ERR"
            End If
        End If
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("CalcolaNOxNH3Istantaneo: " & Error(Err), 1, "OPC")

End Sub

'luca maggio 2018
Public Sub CalcolaNOxNH3MediaOrariaInCorso()

    Dim ValoreNOx As Double
    Dim ValoreNH3 As Double
    Dim FlagNOx As String
    Dim FlagNH3 As String
    Dim ValoreNOxNH3 As Double
    Dim TipoDato As Integer
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri necessari ****
    If IngressoNOXNH3 <= 0 Then
        Call WindasLog("CalcolaNOxNH3MediaOrariaInCorso: IngressoNOXNH3 non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngressoNOX <= 0) Or (IngressoNH3 <= 0) Then
        MediaOraInCorso(0, IngressoNOXNH3) = -9999
        StatusMediaOraInCorso(0, IngressoNOXNH3) = "ERR"
        MediaOraInCorso(1, IngressoNOXNH3) = -9999
        StatusMediaOraInCorso(1, IngressoNOXNH3) = "ERR"
        
        Call WindasLog("CalcolaNOxNH3MediaOrariaInCorso: IngressoNOX o IngressoNH3 non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    For TipoDato = 0 To 1
        ValoreNOx = MediaOraInCorso(TipoDato, IngressoNOX)
        ValoreNH3 = MediaOraInCorso(TipoDato, IngressoNH3)
        FlagNOx = StatusMediaOraInCorso(TipoDato, IngressoNOX)
        FlagNH3 = StatusMediaOraInCorso(TipoDato, IngressoNH3)
        
        If (ValoreNOx <> -9999) And (ValoreNH3 <> -9999) And (InStr("VAL AUX", FlagNOx) > 0) And (InStr("VAL AUX", FlagNH3) > 0) Then
            ValoreNOxNH3 = MediaOraInCorso(TipoDato, IngressoNOX) + (MediaOraInCorso(TipoDato, IngressoNH3) * 2.7)
            
            MediaOraInCorso(TipoDato, IngressoNOXNH3) = ValoreNOxNH3
            StatusMediaOraInCorso(TipoDato, IngressoNOXNH3) = "VAL"
        Else
            MediaOraInCorso(TipoDato, IngressoNOXNH3) = -9999
            StatusMediaOraInCorso(TipoDato, IngressoNOXNH3) = "ERR"
        End If
    Next TipoDato
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("CalcolaNOxNH3Istantaneo: " & Error(Err), 1, "OPC")

End Sub
'Federica settembre 2017 - Calcolo flusso
Public Sub CalcolaFlusso(ByVal parametro As Integer, ByVal parametroFlusso As Integer)

    Dim ValorePortata As Double
    Dim ValoreInquinante As Double
    Dim FlagPortata As String
    Dim FlagInquinante As String
    Dim Flusso As Double
    
    On Error GoTo Gesterrore
    
    If ParametriStrumenti(parametroFlusso).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(parametroFlusso).Acquisizione Then
            ValorePortata = ValIst(1, IngressoPortata)
            ValoreInquinante = ValIst(1, parametro)
            FlagPortata = Status(1, IngressoPortata)
            FlagInquinante = Status(1, parametro)
            
            If (ValorePortata <> -9999) And (ValoreInquinante <> -9999) And (InStr("VAL AUX", FlagPortata) > 0) _
                And (InStr("VAL AUX", FlagInquinante) > 0) Then
                    If (ValoreInquinante <> 0) And (CDbl(Generiche(iDivisorePerFlussi).Par) <> 0) Then
                        'Federica ottobre 2017 - Parametrizzato divisore flusso
                        Flusso = ValorePortata * ValoreInquinante / CDbl(Generiche(iDivisorePerFlussi).Par)
                    Else
                        Flusso = 0
                    End If
                    If Flusso < 0 Then Flusso = 0
                    
                    ValIst(0, parametroFlusso) = Flusso
                    ValIst(1, parametroFlusso) = ValIst(0, parametroFlusso)
                    
                    Status(0, parametroFlusso) = "VAL"
                    Status(1, parametroFlusso) = "VAL"
            Else
                ValIst(0, parametroFlusso) = -9999
                ValIst(1, parametroFlusso) = ValIst(0, parametroFlusso)
                
                Status(0, parametroFlusso) = "ERR"
                Status(1, parametroFlusso) = "ERR"
            End If
        End If
    End If
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("CalcolaFlusso: " & Error(Err), 1, "OPC")

End Sub

'Federica ottobre 2017 - Calcolo Portata da velocità
Public Sub CalcolaPortataDaVelocita()
    
    Dim ValoreVelocita As Double
    Dim FlagVelocita As String
    Dim Raggio As Double
    Dim ValorePortata As Double
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri necessari ****
    If IngressoPortata <= 0 Then
        Call WindasLog("CalcolaPortataDaVelocita: IngressoPortata non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngressoVelocita <= 0) Or (CDbl(Generiche(iRaggioCamino).Par) <= 0) Then
        ValIst(0, IngressoPortata) = -9999
        Status(0, IngressoPortata) = "ERR"
        ValIst(1, IngressoPortata) = -9999
        Status(1, IngressoPortata) = "ERR"
        
        Call WindasLog("CalcolaPortataDaVelocita: IngressoVelocità o diametro non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    '**** Calcolo ****
    If ParametriStrumenti(IngressoPortata).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(IngressoPortata).Acquisizione Then
            ValoreVelocita = ValIst(0, IngressoVelocita)
            FlagVelocita = Status(0, IngressoVelocita)
            Raggio = CDbl(Generiche(iRaggioCamino).Par)
            
            If (ValoreVelocita <> -9999) And (Valido(FlagVelocita)) And (Raggio > 0) Then
                ValorePortata = ValoreVelocita * (Raggio * Raggio * 3.14) * 3600
                
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
    Call WindasLog("CalcolaPortataDaVelocita: " & Error(Err), 1, "OPC")

End Sub

'Federica gennaio 2018 - Calcolo NOX
Public Sub CalcolaNOX(ByVal IngNO As Integer, ByVal IngNO2 As Integer, ByVal IngNOX As Integer)
    
    Dim ValoreNO As Double
    Dim ValoreNO2 As Double
    Dim ValoreNOx As Double
    Dim FlagNO As String
    Dim FlagNO2 As String
    
    On Error GoTo Gesterrore
    
    '**** Verifica parametri ****
    If IngNOX < 0 Then
        Call WindasLog("CalcolaNOX: Ingresso NOX non configurato!", 1, "OPC")
        Exit Sub
    End If
    If (IngNO < 0) Or (IngNO2 < 0) Then
        Call WindasLog("CalcolaNOX: Ingresso NO o NO2 non configurati!", 1, "OPC")
        Exit Sub
    End If
    
    '**** Calcolo ****
    If ParametriStrumenti(IngNOX).TipoAcquisizione = TipiAcquisizione.CALCOLATO Then
        If ParametriStrumenti(IngNOX).Acquisizione Then
            ValoreNO = ValIst(0, IngNO)
            FlagNO = Status(0, IngNO)
            ValoreNO2 = ValIst(0, IngNO2)
            FlagNO2 = Status(0, IngNO2)
            
            If (ValoreNO <> -9999) And (ValoreNO2 <> -9999) And Valido(FlagNO) And Valido(FlagNO2) Then
                ValoreNOx = (ValoreNO * 1.53) + ValoreNO2
                
                ValIst(0, IngNOX) = ValoreNOx
                ValIst(1, IngNOX) = ValIst(0, IngNOX)
                Status(0, IngNOX) = "VAL"
                Status(1, IngNOX) = "VAL"
            Else
                ValIst(0, IngNOX) = 0
                ValIst(1, IngNOX) = ValIst(0, IngNOX)
                Status(0, IngNOX) = "ERR"
                Status(1, IngNOX) = "ERR"
            End If
        End If
    End If
    
    Exit Sub
Gesterrore:
    Call WindasLog("CalcolaNOX: " & Error(Err()), 1, "OPC")
End Sub

