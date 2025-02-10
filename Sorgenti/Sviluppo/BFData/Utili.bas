Attribute VB_Name = "Utili"
Option Explicit

Sub InserisciDecimali(rdValore As Double, ByVal vnNroDecimali As Integer)

    Select Case vnNroDecimali
        
        Case 0
            rdValore = Val(Replace(Format(rdValore, "0"), ",", "."))
        Case 1
            rdValore = Val(Replace(Format(rdValore, "0.0"), ",", "."))
        Case 2
            rdValore = Val(Replace(Format(rdValore, "0.00"), ",", "."))
        Case 3
            rdValore = Val(Replace(Format(rdValore, "0.000"), ",", "."))
        Case 4
            rdValore = Val(Replace(Format(rdValore, "0.0000"), ",", "."))
        Case Else
            rdValore = Val(Replace(Format(rdValore, "0.00000"), ",", "."))
            
    End Select
    
End Sub

'luca luglio 2017
Function DataFineTrimestre(mese As Integer) As Date

    On Error GoTo GestErrore
    
    Select Case mese
        Case 1, 2, 3
            DataFineTrimestre = DateSerial(year(Now), 3, 31)
        Case 4, 5, 6
            DataFineTrimestre = DateSerial(year(Now), 6, 30)
        Case 7, 8, 9
            DataFineTrimestre = DateSerial(year(Now), 9, 30)
        Case 10, 11, 12
            DataFineTrimestre = DateSerial(year(Now), 12, 31)
    End Select
    
    Exit Function

GestErrore:
    Call WindasLog("DataFineTrimestre: " + Error(Err), 1)
    
End Function

'luca luglio 2017
Function DataInizioTrimestre(mese As Integer) As Date

    On Error GoTo GestErrore
    
    Select Case mese
        Case 1, 2, 3
            DataInizioTrimestre = DateSerial(year(Now), 1, 1)
        Case 4, 5, 6
            DataInizioTrimestre = DateSerial(year(Now), 4, 1)
        Case 7, 8, 9
            DataInizioTrimestre = DateSerial(year(Now), 7, 1)
        Case 10, 11, 12
            DataInizioTrimestre = DateSerial(year(Now), 10, 1)
    End Select
    
    Exit Function

GestErrore:
    Call WindasLog("DataInizioTrimestre: " + Error(Err), 1)
    
End Function
