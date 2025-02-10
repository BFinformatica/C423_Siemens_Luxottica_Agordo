Attribute VB_Name = "Flussi"
Option Explicit

Private Function CalcolaFlusso(ByRef rsDati, ByVal CodPar As String, ByVal strStatiImpianto, ByVal DataDa As String, ByVal DataA As String) As Double

    Dim strSQL As String

    On Error GoTo GestErrore
    
    CalcolaFlusso = -9999
    strSQL = "SELECT COUNT(DT_FM) AS CONTEGGIO, SUM(DT_FM) AS DATO FROM " & Tabella & " WHERE DT_STATIONCODE = '" & StationCode & "' " & _
             "AND DT_DATE BETWEEN '" & DataDa & "' AND '" & DataA & "' AND DT_MEASURECOD = '" & CodPar & "' AND DT_CUSTOM1 IN (" & strStatiImpianto & ") " & _
             "AND DT_VALIDFLAG IN (" & strValidValidflags & ") AND DT_FM <> -9999"
    
    If strSQL <> "" Then rsDati.ExecuteSQL (strSQL)
    If Not rsDati.iseof Then
        If rsDati.GetValue("CONTEGGIO") > 0 Then CalcolaFlusso = rsDati.GetValue("DATO")
    End If
            
    Exit Function
    
GestErrore:
    Call WindasLog("CalcolaFlusso: " & Error(Err()), 1)
    
End Function

Public Function CalcolaFM(ByRef rsDati, ByVal CodPar As String, ByVal strStatiImpianto, ByVal DataDa As String, ByVal DataA As String) As Double

    Dim ParametroPortata As String
    
    On Error GoTo GestErrore
    
    CalcolaFM = -9999
    
    'luca luglio 2017
    If IngressoQFUMI >= 0 Then
        ParametroPortata = SuperTrim(gaConfigurazioneArchivio(IngressoQFUMI).STRUM.NomeParametro)
        'Portata
        If CalcolaFlusso(rsDati, ParametroPortata, strStatiImpianto, DataDa, DataA) = -9999 Then Exit Function
        CalcolaFM = CalcolaFlusso(rsDati, CodPar, strStatiImpianto, DataDa, DataA)
    End If
    
    Exit Function
    
GestErrore:
    Call WindasLog("CalcolaFM: " & Error(Err()), 1)

End Function
