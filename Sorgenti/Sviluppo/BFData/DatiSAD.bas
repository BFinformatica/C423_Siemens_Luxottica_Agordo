Attribute VB_Name = "DatiSAD"
Option Explicit

Sub DatiSADSalvaSuFile(NomeFile As String, Elabdate As Date)

    Dim iIdx As Integer
    Dim iIdx1 As Integer 'Nicolò gennaio 2018
    Dim DataFileARPA As String
    Dim fso, f, OraFileARPA, valore
    Dim idDb As Double
    
    On Error GoTo GestErrore:
    
    'Nicolò Gennaio 2018 sostituisco tutti i CodiceMonitor con CodiceMonitorItantaneiTq al fine di poter creare degi sad di output con parametri in meno e codici monitor differenti rispetto al sad originale
    
    'Nicolò gennaio 2018 ****************************
    Dim IndexesForADIADM() As Integer
    Call FillIndexesForADIADM(IndexesForADIADM)
    '************************************************
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Dir(NomeFile) <> "" Then Kill NomeFile
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
    
    
    
    'Scrivo intestazione e apro for writing
    Set f = fso.OpenTextFile(NomeFile, ForWriting, True)

    '***** Riga 1 Id. del software utilizzato dal gestore *****
    f.Writeline Nome_Software_4343

    '***** Riga 2 Codice Impianto assegnato da ARPA *****
    f.Writeline Nome_Impianto_4343

    '***** Riga 3 Codice Monitor *****
    f.Write "#" & String(10, " ")
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
        iIdx = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorIst_TQ) <> "" Then
            f.Write Chr(9) & Chr(9) & gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorIst_TQ
        End If
    Next iIdx1
    f.Writeline Chr(9)

    '***** Riga 4 Unità di misura dei Codici Monitor *****
    f.Write "#" & String(10, " ")
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
        iIdx = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorIst_TQ) <> "" Then
            f.Write Chr(9) & Chr(9) & Trim(gaConfigurazioneArchivio(iIdx).STRUM.UnitaMisuraTq) 'Nicolò Luglio 2016 passo a unitamisuratq invece di unitamisura
        End If
    Next iIdx1
    f.Writeline Chr(9)
    
    'Scorro tutti i dati della matrice
    DataFileARPA = Mid(CStr(Elabdate), 7, 4) & Mid(CStr(Elabdate), 4, 2) & Mid(CStr(Elabdate), 1, 2)
    Dim i As Long
    For i = 0 To 17279
        Dim tempdate As Date
        tempdate = DateAdd("s", i * 5, DateTimeSerial(0, 1, 1, 0, 0, 0))
        OraFileARPA = Format(tempdate, "hh.nn.ss")
        
        '***** Riga >=5 Data, Ora, valore istantaneo, stato della misura *****
        f.Write DataFileARPA & Chr(9) & OraFileARPA
        For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
            iIdx = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
            If SuperTrim(gaConfigurazioneArchivio(iIdx).STRUM.CodiceMonitorIst_TQ) <> "" Then
                idDb = CodParametro(CDbl(gaConfigurazioneArchivio(iIdx).STRUM.iddatabase))
                If Dati(i, idDb, 0) = "" Then
                    valore = "---"
                Else
                    valore = Trim(Dati(i, idDb, 0))
                End If
                f.Write Chr(9) & valore & Chr(9) & Trim(Dati(i, idDb, 1))
            End If
        Next iIdx1
        f.Writeline Chr(9)
    Next i
    
    f.Close
    Set f = Nothing
    Set fso = Nothing
   
    Exit Sub
GestErrore:
    Call WindasLog("BFdata DatiSADSalvaFile: " + Error(Err), 1)
    Resume Next

End Sub

Private Sub FillIndexesForADIADM(ByRef IndexesForADIADM() As Integer)
    'Nicolò gennaio 2017
    'Popola e ridimensiona l'array inserendovi solamente gli indici per accedere alla configurazione dei parametri nell'ordine impostato per l'output ADI e ADM
    'Maggiore è il valore dell'ordine prima viene scritta la misura.
    On Error GoTo GestErrore
    
    'Dimensiono
    ReDim IndexesForADIADM(gnNroParametriStrumenti)
    
    
    Dim A() As Double
    Dim B() As Double
    Dim Bindex As Integer
    
    ReDim A(gnNroParametriStrumenti, 1) As Double
    ReDim B(gnNroParametriStrumenti, 1) As Double
    Bindex = 0
    
    Dim i As Integer
    For i = 0 To gnNroParametriStrumenti
        A(i, 0) = i
        A(i, 1) = gaConfigurazioneArchivio(i).STRUM.OrdineScritturaADIADM
        'Debug.Print (A(i, 0) & " - " & gaConfigurazioneArchivio(i).STRUM.NomeParametro & " - " & A(i, 1))
    Next i
    
    Dim AUpperLimit As Integer
    AUpperLimit = UBound(A)
    While AUpperLimit >= 0
        'TROVO MAX
        Dim maxPosition As Integer
        Dim maxValue As Double
        maxValue = -10000
        For i = 0 To AUpperLimit
            If A(i, 1) > maxValue Then
                maxPosition = i
                maxValue = A(i, 1)
            End If
        Next i
        
        'COPIO MAX IN B
        B(Bindex, 0) = A(maxPosition, 0)
        B(Bindex, 1) = A(maxPosition, 1)
        Bindex = Bindex + 1
        
        'RIMUOO ELEMENTO MAX DA A
        Dim ForLoopCounter As Integer
        Dim UpperLimitsOfArray As Integer, LowerLimitsOfArray As Integer
        UpperLimitsOfArray = AUpperLimit 'UBound(A)
        For ForLoopCounter = maxPosition To UpperLimitsOfArray - 1
            A(ForLoopCounter, 0) = A(ForLoopCounter + 1, 0)
            A(ForLoopCounter, 1) = A(ForLoopCounter + 1, 1)
        Next ForLoopCounter
        'Siccome il redim preserve non va, invece di rimuovere compatto solamente i dati all'inizio dell'array e limito la successiva ricerca del MAX
        'LowerLimitsOfArray = LBound(A)
        'ReDim Preserve A(UpperLimitsOfArray - 1, 1) As Double
        AUpperLimit = AUpperLimit - 1
        
    Wend
    
'    Debug.Print ("ORDERED!!")
    
    For i = 0 To UBound(B)
        IndexesForADIADM(i) = B(i, 0)
'        Debug.Print (IndexesForADIADM(i) & " - " & gaConfigurazioneArchivio(IndexesForADIADM(i)).STRUM.NomeParametro & " - " & gaConfigurazioneArchivio(IndexesForADIADM(i)).STRUM.OrdineScritturaADIADM)
    Next i
    
    
    Exit Sub
GestErrore:
    Call WindasLog("FillIndexesForADIADM: " & Err.Description, 1)
End Sub

Sub DatiSADSalva(Elabdate As Date)

    Dim iIdx As Integer
    Dim iOra As Integer
    Dim iSecondi As Integer
    Dim NomeFile As String
    Dim DataFileARPA As String

    On Error GoTo GestErrore
    
    DataFileARPA = Format(DateSerial(year(Elabdate), month(Elabdate), day(Elabdate)), "yyyymmdd")
    
    
    NomeFile = gsDirLavoro & "Windas03" & CStr(NumeroLineaBFData) & "\" & PathARPA_FileUnico & "\"
    'Creo la cartella
    If Dir(NomeFile) = "" Then
        MkDir NomeFile
    End If
    'Alby Luglio 2016
    NomeFile = NomeFile & Nome_File_4343 & "_" & DataFileARPA & ".SAD"
    If Dir(NomeFile) <> "" Then Kill CStr(NomeFile)
    
    'Carico tutto in un array
    For iOra = 0 To 23
        For iSecondi = 0 To 719
            For iIdx = 0 To gnNroParametriStrumenti
            
                Dati(iSecondi + (720 * iOra), iIdx, 0) = FormattaNumero(Valore_5_Secondi(iOra, iIdx, iSecondi), -2)
                Dati(iSecondi + (720 * iOra), iIdx, 1) = Status_5_Secondi(iOra, iIdx, iSecondi)
            
            Next iIdx
        Next iSecondi
    Next iOra
    
    Call DatiSADSalvaSuFile(NomeFile, Elabdate)
    
    Exit Sub

GestErrore:
    Call WindasLog("BFdata DatiSADSalva: " + Error(Err), 1)

End Sub

Sub ElaboraSalvaDatiConcludoADM(Elabdate As Date, numeromedie() As Integer)

    Dim nn As Integer
    Dim nl As Integer
    Dim Ora As Integer
    Dim nFile As Integer
    Dim NomeFile As String
    Dim nmedie As Integer
    Dim NumeroLinea As Integer
    Dim CodiceMonitor As String
    Dim NumImp As Integer
    Dim Numd As Integer
    Dim MediaOraStr(2) As String
    Dim MassimoOraStr(2) As String
    Dim MinimoOraStr(2) As String
    Dim StdDevStr(2) As String
    Dim CheckDate As Date

    On Local Error GoTo GestErrore

    If UCase(Tabella) = "WDS_HALF" Then
        NomeFile = Nome_File_4343 & "_" & Format(Elabdate, "yyyymmdd") & ".1800.MEDIE"
    ElseIf UCase(Tabella) = "WDS_10MINCO" Then
        NomeFile = Nome_File_4343 & "_" & Format(Elabdate, "yyyymmdd") & ".600.MEDIE"
    Else
        NomeFile = Nome_File_4343 & "_" & Format(Elabdate, "yyyymmdd") & ".3600.MEDIE"
    End If
    NomeFile = gsDirLavoro & "Windas03" & CStr(NumeroLineaBFData) & "\" & PathARPA_FileUnico & "\" & NomeFile

    '***** Creazione cartella per file ARPA *****
    If (Dir(gsDirLavoro & "Windas03" & CStr(NumeroLineaBFData) & "\" & PathARPA_FileUnico, vbDirectory) = "") Then
        MkDir (gsDirLavoro & "Windas03" & CStr(NumeroLineaBFData) & "\" & PathARPA_FileUnico)
    End If

    Form1.Label1.Caption = "Elaborazione dei parametri..."
    'DoEvents

    'Nicolò gennaio 2018 ****************************
    Dim IndexesForADIADM() As Integer
    Call FillIndexesForADIADM(IndexesForADIADM)
    Dim iIdx1
    Dim i
    Dim unMisura
    '************************************************
    'Data
    CheckDate = Elabdate

    'Apertura file per scrittura
    nFile = FreeFile
    Open NomeFile For Output As #nFile

    '***** Scrittura su file ****
    Print #nFile, Nome_Software_4343
    Print #nFile, Nome_Impianto_4343

    '***** descrizione parametro *****
    Print #nFile, Chr(9) & Chr(9) & Chr(9) & Chr(9);
    '***** ciclo per riportare i codici monitor delle medie tal quali
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
    i = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If (Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_TQ) <> "") _
              And gaConfigurazioneArchivio(i).STRUM.iddatabase <> 670 Then
            Print #nFile, Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_TQ) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9);
        End If
    Next iIdx1
    '***** ciclo per riportare i codici monitor delle medie elaborate
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
    i = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If (Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_EL) <> "") _
              And (gaConfigurazioneArchivio(i).STRUM.iddatabase <> 670) Then
            Print #nFile, Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_EL) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9);
        End If
    Next iIdx1
    
    Print #nFile, "stato_30" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_31" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_32" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_33" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_34" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_35" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_36" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_37" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "stato_38" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) ';

    '***** Unità di misura *****
    Print #nFile, Chr(9) & Chr(9) & Chr(9) & Chr(9);
    '***** ciclo per riportare le unità di misura delle medie tal quali
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
    i = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If (Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_TQ) <> "") _
              And (gaConfigurazioneArchivio(i).STRUM.iddatabase <> 670) Then
            unMisura = Trim(gaConfigurazioneArchivio(i).STRUM.UnitaMisuraTq)
            Print #nFile, unMisura & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9);
        End If
    Next iIdx1
    '***** ciclo per riportare le unità di misura  delle medie elaborate
    For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
    i = IndexesForADIADM(iIdx1) 'Nicolò gennaio 2018
        If (Trim(gaConfigurazioneArchivio(i).STRUM.CodiceMonitorMed_EL) <> "") _
              And (gaConfigurazioneArchivio(i).STRUM.iddatabase <> 670) Then
            unMisura = Trim(gaConfigurazioneArchivio(i).STRUM.UnitaMisura)
            Print #nFile, unMisura & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9);
        End If
    Next iIdx1
    Print #nFile, "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                  "---" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) ';

    '###################### ############################## ##################################
    For Ora = 0 To 23
        For nmedie = 1 To numeromedie(Ora)
            Select Case TipoMedia
                Case 0
                    '***** media mezzora *****
                    If nmedie = 1 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".00.00" & Chr(9) & Trim(Str(Ora)) & ".29.59" & Chr(9) & "360" & Chr(9);
                    ElseIf nmedie = 2 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".30.00" & Chr(9) & Trim(Str(Ora)) & ".59.59" & Chr(9) & "360" & Chr(9);
                    End If

                Case 1
                    '***** media 10 minuti *****
                    If nmedie = 1 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".00.00" & Chr(9) & Trim(Str(Ora)) & ".09.59" & Chr(9) & "120" & Chr(9);
                    ElseIf nmedie = 2 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".10.00" & Chr(9) & Trim(Str(Ora)) & ".19.59" & Chr(9) & "120" & Chr(9);
                    ElseIf nmedie = 3 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".20.00" & Chr(9) & Trim(Str(Ora)) & ".29.59" & Chr(9) & "120" & Chr(9);
                    ElseIf nmedie = 4 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".30.00" & Chr(9) & Trim(Str(Ora)) & ".39.59" & Chr(9) & "120" & Chr(9);
                    ElseIf nmedie = 5 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".40.00" & Chr(9) & Trim(Str(Ora)) & ".49.59" & Chr(9) & "120" & Chr(9);
                    ElseIf nmedie = 6 Then
                        Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Trim(Str(Ora)) & ".50.00" & Chr(9) & Trim(Str(Ora)) & ".59.59" & Chr(9) & "120" & Chr(9);
                    End If

                Case 2
                    '***** media oraria *****
                    Print #nFile, Format(Elabdate, "yyyymmdd") & Chr(9) & Format(Ora, "00") & ".00.00" & Chr(9) & Format(Ora, "00") & ".59.59" & Chr(9) & "720" & Chr(9);

            End Select

            For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
                nn = IndexesForADIADM(iIdx1)
                For Numd = 0 To 2
                    'Alby Ottobre 2013 da verificare nr. decimali scritti in ADM

                    If MedieOra(Ora, nn, Numd, nmedie) <> -9999 Then
                        MediaOraStr(Numd) = FormattaNumero(MedieOra(Ora, nn, Numd, nmedie), -2)
                    Else
                        MediaOraStr(Numd) = "---"
                    End If

                    If StdDev(Ora, nn, Numd, nmedie) <> -9999 Then
                        StdDevStr(Numd) = FormattaNumero((StdDev(Ora, nn, Numd, nmedie)), -2)
                    Else
                        StdDevStr(Numd) = "---"
                    End If

                    If minimo(Ora, nn, 0, nmedie) <> 999999999 Then
                        MinimoOraStr(Numd) = FormattaNumero((minimo(Ora, nn, Numd, nmedie)), -2)
                    Else
                        MinimoOraStr(Numd) = "---"
                    End If

                    If massimo(Ora, iIdx1, 0, nmedie) <> -999999999 Then
                        MassimoOraStr(Numd) = FormattaNumero((massimo(Ora, nn, Numd, nmedie)), -2)
                    Else
                        MassimoOraStr(Numd) = "---"
                    End If

                Next Numd

                '***** media oraria e semioraria *****
                If Len(Trim(gaConfigurazioneArchivio(nn).STRUM.CodiceMonitorMed_TQ)) > 0 Then
                    If ContaTutti_5_secondi(Ora, nn, 0, nmedie) = 0 Then
                        '***** nessun dato *****
                        Print #nFile, "0" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "ERR" & Chr(9);
                    Else
                        Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, nn, 0, nmedie))) & Chr(9) & Trim(Str(ContaOraOK(Ora, nn, 0, nmedie))) & Chr(9) & MediaOraStr(0) & Chr(9) & MinimoOraStr(0) & Chr(9) & MassimoOraStr(0) & Chr(9) & StdDevStr(0) & Chr(9) & Trim(StsMedieOra(Ora, nn, 0, nmedie)) & Chr(9);
                    End If
                End If

            Next iIdx1

            For iIdx1 = 0 To UBound(IndexesForADIADM) 'Nicolò gennaio 2018 non scorro più fino a gnNroParametriStrumenti ma processo solo i parametri che devo scrivere come outpt ADI ADM
                nn = IndexesForADIADM(iIdx1)
                For Numd = 0 To 2
                    'Alby Ottobre 2013 da verificare nr. decimali scritti in ADM

                    If MedieOra(Ora, nn, Numd, nmedie) <> -9999 Then
                        MediaOraStr(Numd) = FormattaNumero(MedieOra(Ora, nn, Numd, nmedie), -2)
                    Else
                        MediaOraStr(Numd) = "---"
                    End If

                    If StdDev(Ora, nn, Numd, nmedie) <> -9999 Then
                        StdDevStr(Numd) = FormattaNumero((StdDev(Ora, nn, Numd, nmedie)), -2)
                    Else
                        StdDevStr(Numd) = "---"
                    End If

                    If minimo(Ora, nn, 0, nmedie) <> 999999999 Then
                        MinimoOraStr(Numd) = FormattaNumero((minimo(Ora, nn, Numd, nmedie)), -2)
                    Else
                        MinimoOraStr(Numd) = "---"
                    End If

                    If massimo(Ora, nn, 0, nmedie) <> -999999999 Then
                        MassimoOraStr(Numd) = FormattaNumero((massimo(Ora, nn, Numd, nmedie)), -2)
                    Else
                        MassimoOraStr(Numd) = "---"
                    End If

                Next Numd

                If Len(Trim(gaConfigurazioneArchivio(nn).STRUM.CodiceMonitorMed_EL)) > 0 Then
                    If ContaTutti_5_secondi(Ora, nn, 0, nmedie) = 0 Then
                        '***** nessun dato *****
                        Print #nFile, "0" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "ERR" & Chr(9);
                    Else
                        Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, nn, 0, nmedie))) & Chr(9) & Trim(Str(ContaOraOK(Ora, nn, 1, nmedie))) & Chr(9) & MediaOraStr(1) & Chr(9) & MinimoOraStr(1) & Chr(9) & MassimoOraStr(1) & Chr(9) & StdDevStr(1) & Chr(9) & Trim(StsMedieOra(Ora, nn, 1, nmedie)) & Chr(9);
                    End If
                End If
            Next iIdx1
            
            '***** Stati Impianto *****
            '***** stato 30 impianto funzionante *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercRegime(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);

            '***** stato 31 impianto in accensione (minimo tecnico) *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercMinTec(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);

            '***** stato 32 impianto in spegnimento *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercSpegnimento(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);

            '***** stato 33 impianto in manutenzione *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercManutenzione(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);

            '***** stato 34 impianto fermo *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercFermo(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);

            '***** stato 35 impianto in guasto *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercGuasto(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);
            
            '***** stato 36 impianto in anomalia *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercAnomalo(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);
            
            '***** stato 37 taratura del misuratore di polveri *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercPolveri(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);
            
            '***** stato 38 altro *****
            Print #nFile, Trim(Str(ContaTutti_5_secondi(Ora, IngressoIMPIANTO, 0, nmedie))) & Chr(9) & Trim(Str(PercAltro(Ora, nmedie))) & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9) & "---" & Chr(9);       'michele ottobre 2013: aggiunto ";" finale
        Next nmedie

'        'michele ottobre 2013: colonna aggiuntiva con l'O2 di riferimento
        'Print #nFile, "720" & Chr(9) & "720" & Chr(9) & O2riferimento & Chr(9) & "0.000" & Chr(9) & "0.000"; Chr(9) & "0.000"; Chr(9) & "VAL"
        Print #nFile, ""
    Next Ora

    Close #nFile

   Exit Sub

GestErrore:
    Call WindasLog("ElaboraSalvaDatiConcludoADM: " + Error(Err), 1)
    Resume Next
    
End Sub


