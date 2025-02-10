Attribute VB_Name = "Elementari"
Option Explicit

Sub CaricaDatidaDB(Elabdate As Date, ByVal iConnessione As Integer)

    Dim rsDati As Object
    Dim NomeDBElementare As String
    Dim NomeTabella As String
    Dim strSQL As String
    Dim Station As String
    Dim iIdx As Integer
    Dim CRLF As String * 2
    Dim Ora As Integer
    Dim Minuto As Integer
    Dim Secondo As Integer
    Dim cinque_secondi As Integer

    'Alby Ottobre 2016
    On Error GoTo GestErrore

    'Call GetConnectionParam
    NewDataObj rsDati, iConnessione

    CRLF = Chr(13) + Chr(10)
    Station = connDB(iConnessione).StationCode
    'luca giugno 2017
    'NomeDBElementare = Station & "_" + Format(Elabdate, "yyyy")
    NomeDBElementare = Station & "_" + Format(Elabdate, "yyyymm")
    NomeTabella = "BFM" + Format(Elabdate, "yyyymmdd")
    strSQL = "USE [" & NomeDBElementare & "]" + CRLF
    rsDati.ExecuteSQL strSQL
    
    strSQL = "SELECT * FROM " + NomeTabella + " WHERE DT_STATIONCODE='" + Station + "'"
    rsDati.SelectionFast strSQL
    
    For iIdx = 0 To gnNroParametriStrumenti
        
        rsDati.m_filter = "DT_MEASURECOD='" + Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro) + "'"
        rsDati.movefirst
        
        Do While Not rsDati.iseof
            Ora = Mid(rsDati.GetValue("DT_DATETIME"), 9, 2)
            Minuto = Mid(rsDati.GetValue("DT_DATETIME"), 11, 2)
            Secondo = Mid(rsDati.GetValue("DT_DATETIME"), 13, 2)
            cinque_secondi = (12 * Minuto) + (Secondo / 5)
            ContaTuttiSecondiMediaOra(Ora, iIdx, cinque_secondi) = 1
            Form1.Label1.Caption = "Lettura dati elementari da Server: " + connDB(iConnessione).AppServer + " " + Trim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro) + " Ore " + Format(Ora, "00") + " " + Format(Minuto, "00") + " " + Format(Secondo, "00")
            DoEvents
            
            'luca maggio 2017
            'If UCase(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro)) = "IMP_L" & right1(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro), 1) Then
            If UCase(SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.NomeParametro)) = "IMPIANTO" Then
                If Valore_5_Secondi(Ora, iIdx, cinque_secondi) = -9999 And rsDati.GetValue("DT_VALUE") <> -9999 Then
                    Valore_5_Secondi(Ora, iIdx, cinque_secondi) = rsDati.GetValue("DT_VALUE")
                End If
            Else
                If (Status_5_Secondi(Ora, iIdx, cinque_secondi) <> "VAL" And rsDati.GetValue("DT_VALIDFLAG") = "VAL") Or Status_5_Secondi(Ora, iIdx, cinque_secondi) = "-9999" Then
                    Valore_5_Secondi(Ora, iIdx, cinque_secondi) = rsDati.GetValue("DT_VALUE")
                    Status_5_Secondi(Ora, iIdx, cinque_secondi) = rsDati.GetValue("DT_VALIDFLAG")
                    'dati normalizzati = dati tal quali
                    Valore_5_Secondi_N(Ora, iIdx, cinque_secondi) = Valore_5_Secondi(Ora, iIdx, cinque_secondi)
                    Status_5_Secondi_N(Ora, iIdx, cinque_secondi) = Status_5_Secondi(Ora, iIdx, cinque_secondi)
                End If
            End If
            rsDati.MoveNext
        Loop
    Next

    strSQL = "USE [" & connDB(iConnessione).AppDatabase & "]" + CRLF
    rsDati.ExecuteSQL strSQL
    Set rsDati = Nothing
    Form1.Label1.Caption = "Elaborazione dati..."
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata CaricaDatidaDB: " + Error(Err), 1)
    Resume Next
End Sub

Sub ElaboraCaricaDatiElementari(Elabdate As Date, modo)
    
    Dim i As Integer
    
    On Error GoTo GestErrore
    
    'Nicola 29/11/216
    'Carico i dati sad da db per tutte le connessioni configurate
    'in base all'ordine di caricamento configurato sul connection.xml
    
    For i = 0 To UBound(connDB)
        Call CaricaDatidaDB(Elabdate, i)
    Next
    
    Exit Sub
GestErrore:
    Call WindasLog("BFdata ElaboraCaricaDatiElementari: " + Error(Err), 1)
    Resume Next
End Sub

