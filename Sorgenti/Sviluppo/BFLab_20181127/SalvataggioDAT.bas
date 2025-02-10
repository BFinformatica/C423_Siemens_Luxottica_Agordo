Attribute VB_Name = "SalvataggioDAT"
Option Explicit

Dim DatiAcquisitiDAT(11)
Dim SommatoriaMediaOraDaSecondi(2, 72)
Dim ContaSecondiMediaOra(2, 72)
Dim ContaSecondiMediaOraTotale(2, 72) As Integer    'luca 06/09/2016 Conta totale per calcolo ID

Sub SalvaDAT()

    Dim iIndice As Integer

    On Error GoTo Gesterrore

    'daniele luglio 2013 bolgiano: aggiungo salvataggio nuovi sad
    For iIndice = 0 To 11 '???
        Select Case iIndice
            Case sec_ndx
                If Not DatiAcquisitiDAT(iIndice) Then
                    'Alby Dicembre 2015
                    Call ElaboraProgressivi
                    'luca maggio 2018
                    Call DatiDATMSalvaSuFile(1)
                    Call DatiDATMSalvaSuFile(0) 'Nicolò Aprile 2016 salvo dati elementari tal quali
                End If
                DatiAcquisitiDAT(iIndice) = True
            Case Else
                DatiAcquisitiDAT(iIndice) = False
        End Select
    Next
    
Exit Sub

Gesterrore:
    Call WindasLog("SalvaDAT ", 1, OPC)

End Sub

Sub DatiDATMSalvaSuFile(TipoDato)

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso, f, NomeFile, iIDParametro, DataFileDAT, OraFileDAT, valore
    Dim sec_ndx, sec_ndx_str
    Dim oraSAD
    Dim minSAD
    
    'Alby Ottobre 2013
    Dim Estensione(2)
    
    On Error GoTo Gesterrore

    'Alby Ottobre 2013
    Estensione(0) = ".DATQ" 'Nicolò aprile 2016 salvo dati tq separatamente
    Estensione(1) = ".DATM"
    Estensione(2) = ".DATS"
    
    sec_ndx = second(Now) \ 5
    oraSAD = hour(Now)
    minSAD = minute(Now)
    sec_ndx_str = CStr(sec_ndx * 5)
    DataFileDAT = Format(Now, "yyyymmdd")
    
    'luca giugno 2017 sistemo
    TimeStamp = DateTimeSerial(year(Now), month(Now), day(Now), hour(Now), minute(Now), CInt(Right("00" & sec_ndx_str, 2)))
    
    OraFileDAT = Right("00" & hour(Now), 2) & "." & Right("00" & minute(Now), 2) & "." & Right("00" & sec_ndx_str, 2)
    NomeFile = Nome_File_4343 & "_" & DataFileDAT & Estensione(TipoDato)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Creo la cartella
    If Not (fso.folderexists(PathDAT)) Then
        MkDir PathDAT
    End If
    
    If (fso.FileExists(PathDAT & "\" & CStr(NomeFile))) Then
        Set f = fso.OpenTextFile(PathDAT & "\" & CStr(NomeFile), ForAppending, True)
        '***** Riga >=5 Data, Ora, valore istantaneo, stato della misura *****
        f.Write DataFileDAT & Chr(9) & OraFileDAT
        
        For iIDParametro = 0 To gnNroParametriStrumenti
            If ValIst(TipoDato, iIDParametro) <= -8888 Then
                valore = "---"
            Else
                valore = Trim(Replace(CStr(FormatNumber(ValIst(TipoDato, iIDParametro), 2, -2, -2, 0)), ",", "."))
            End If
            
            f.Write Chr(9) & valore & Chr(9) & Trim(CStr(Status(TipoDato, iIDParametro)))
        Next iIDParametro
        f.Writeline Chr(9)
    Else
        Set f = fso.OpenTextFile(PathDAT & "\" & CStr(NomeFile), ForWriting, True)
        
        '***** Riga 1 Id. del software utilizzato dal gestore *****
        f.Writeline Nome_Software_4343
                                
        '***** Riga 2 Codice Impianto assegnato da ARPA *****
        f.Writeline Nome_Impianto_4343
        
        '***** Riga 3 Codice Monitor *****
        f.Write "#" & String(10, " ")
        For iIDParametro = 0 To gnNroParametriStrumenti
            f.Write Chr(9) & Chr(9) & ParametriStrumenti(iIDParametro).DescrParametro
        Next iIDParametro
        f.Writeline Chr(9)
        
        '***** Riga 4 Unità di misura dei Codici Monitor *****
        f.Write "#" & String(10, " ")
        For iIDParametro = 0 To gnNroParametriStrumenti
            f.Write Chr(9) & Chr(9) & ParametriStrumenti(iIDParametro).UnitaMisura
        Next iIDParametro
        f.Writeline Chr(9)
        
        '***** Riga >=5 Data, Ora, valore istantaneo, stato della misura *****
        f.Write DataFileDAT & Chr(9) & OraFileDAT
        For iIDParametro = 0 To gnNroParametriStrumenti
            If ValIst(TipoDato, iIDParametro) <= -8888 Then
                valore = "---"
            Else
                valore = Trim(Replace(CStr(FormatNumber(ValIst(TipoDato, iIDParametro), 2, -2, -2, 0)), ",", "."))
            End If
            
            f.Write Chr(9) & valore & Chr(9) & Trim(CStr(Status(TipoDato, iIDParametro)))
        Next iIDParametro
        f.Writeline Chr(9)
    End If
    
    f.Close

    Exit Sub
    
Gesterrore:
    Call WindasLog("DatiDATMSalvaSuFile " + Error(Err), 1, OPC)
    Resume Next
End Sub

Private Sub ElaboraProgressivi()
    
    Dim OraAttuale As String
    Dim iIParametro As Integer
    Dim iIndice As Integer
    Static Ora As String
    
    On Error GoTo Gesterrore
    
    'luca aprile 2017
    If OreSemiore = TIPO_MEDIE_ORARIE Then
        OraAttuale = Format(Now, "hh")
    ElseIf OreSemiore = TIPO_MEDIE_SEMIORARIE Then
        If minute(Now) <= 29 Then
            OraAttuale = CStr(TimeSerial(hour(Now), 0, 0))
        Else
            OraAttuale = CStr(TimeSerial(hour(Now), 30, 0))
        End If
    End If
    
    'luca aprile 2017
    '***** AZZERAMENTO *****
    If Ora <> OraAttuale Then
        For iIParametro = 0 To gnNroParametriStrumenti
            For iIndice = 0 To 1
                SommatoriaMediaOraDaSecondi(iIndice, iIParametro) = 0
                ContaSecondiMediaOra(iIndice, iIParametro) = 0
                ContaSecondiMediaOraTotale(iIndice, iIParametro) = 0 'luca 06/09/2016
                MediaOraInCorso(iIndice, iIParametro) = -9999
                StatusMediaOraInCorso(iIndice, iIParametro) = "ERR"
                ID_MediaOraInCorso(iIndice, iIParametro) = 0
            Next iIndice
        Next iIParametro
        Ora = OraAttuale
    End If

    For iIParametro = 0 To gnNroParametriStrumenti
        '***** calcolo media in corso grezza *****
        ContaSecondiMediaOraTotale(0, iIParametro) = ContaSecondiMediaOraTotale(0, iIParametro) + 1
        'luca luglio 2017
        If Valido(Status(0, iIParametro)) Then
            SommatoriaMediaOraDaSecondi(0, iIParametro) = SommatoriaMediaOraDaSecondi(0, iIParametro) + ValIst(0, iIParametro)
            ContaSecondiMediaOra(0, iIParametro) = ContaSecondiMediaOra(0, iIParametro) + 1
            If ContaSecondiMediaOra(0, iIParametro) > 0 Then
                MediaOraInCorso(0, iIParametro) = SommatoriaMediaOraDaSecondi(0, iIParametro) / ContaSecondiMediaOra(0, iIParametro)
            Else
                MediaOraInCorso(0, iIParametro) = -9999
            End If
        End If
        
        '***** ID Media in corso grezza *****
        If ContaSecondiMediaOraTotale(0, iIParametro) > 0 Then
            ID_MediaOraInCorso(0, iIParametro) = ContaSecondiMediaOra(0, iIParametro) / ContaSecondiMediaOraTotale(0, iIParametro) * 100
            If ID_MediaOraInCorso(0, iIParametro) > 100 Then ID_MediaOraInCorso(0, iIParametro) = 100
        Else
            ID_MediaOraInCorso(0, iIParametro) = 0
        End If
        
        '***** VALIDITA' Media in corso grezza *****
        If ID_MediaOraInCorso(0, iIParametro) >= 70 Then
           StatusMediaOraInCorso(0, iIParametro) = "VAL"
        Else
           StatusMediaOraInCorso(0, iIParametro) = "ERR"
        End If
    Next iIParametro
    
    For iIParametro = 0 To gnNroParametriStrumenti
        ContaSecondiMediaOraTotale(1, iIParametro) = ContaSecondiMediaOraTotale(1, iIParametro) + 1
        'luca luglio 2017
        If Valido(Status(1, iIParametro)) Then
            ContaSecondiMediaOra(1, iIParametro) = ContaSecondiMediaOra(1, iIParametro) + 1
        End If
        
        Call ElaborazioniDiLegge(iIParametro, MediaOraInCorso(1, iIParametro), StatusMediaOraInCorso(1, iIParametro), True)
        
        '***** ID Media in corso normalizzata *****
        If ContaSecondiMediaOraTotale(1, iIParametro) > 0 Then
            ID_MediaOraInCorso(1, iIParametro) = ContaSecondiMediaOra(1, iIParametro) / ContaSecondiMediaOraTotale(1, iIParametro) * 100
            If ID_MediaOraInCorso(1, iIParametro) > 100 Then ID_MediaOraInCorso(1, iIParametro) = 100
        Else
            ID_MediaOraInCorso(1, iIParametro) = 0
        End If
        
        '***** VALIDITA' Media in corso normalizzata *****
        If ID_MediaOraInCorso(1, iIParametro) >= 70 Then
           StatusMediaOraInCorso(1, iIParametro) = "VAL"
        Else
           StatusMediaOraInCorso(1, iIParametro) = "ERR"
        End If
    Next iIParametro
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("ElaboraProgressivi " + Error(Err), 1, OPC)

End Sub


