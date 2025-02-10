Attribute VB_Name = "RecuperoDati"
Option Explicit

'Federica agosto 2017 - Procedura di recupero dati (ADAM 5560)
Public Sub RecuperoDatiADAM5560()

    Dim Record As String
    Dim Dati() As String
    Dim Data As Date
    
    Dim iCan As Integer

    'Alby Agosto 2017
    On Error GoTo Gesterrore
    
    Data = Now
    
    iCan = FreeFile
    
    Call WindasLog("RecuperoDatidaADAM5560: Inizio recupero dati", 0, OPC)
    
    'TODO: Attualmente viene recuperato solo OGGI
    
    Open AdamPath + "\" + Format(Data, "yyyymmdd") + ".txt" For Input As #iCan
    
    Do While Not EOF(iCan)
        Line Input #iCan, Record
        
        Record = Replace(Record, Chr(214), "")
        If InStr(Record, Chr(214)) > 0 Then Stop
        
        Dati = Split(Record, Chr(9))
        'Federica settembre 2017 - Segnalo l'importazione come fallita se ci sono stati problemi nel processare la riga
        If Not RecuperoDatiADAM5560ScriviDB(Dati) Then
            Close iCan
            
            Call RecuperoDatiADAM5560Result("FAIL")
            Exit Sub
        End If
    Loop
    
    Close iCan
    
    'Federica settembre 2017 - Recupero concluso con successo
    Call RecuperoDatiADAM5560Result("OK")
    Call WindasLog("RecuperoDatiADAM5560: Fine recupero dati", 0, OPC)

    Exit Sub

Gesterrore:

    Call WindasLog("RecuperoDatiADAM5560 " + Error(Err), 1, OPC)
    'Federica settembre 2017 - Se si verificano errori considero il recupero fallito
    Call RecuperoDatiADAM5560Result("FAIL")

End Sub

'Federica settembre 2017 - Lettura risultato recupero dati da file
Public Sub RecuperoDatiADAM5560ReadFile(ByRef RecuperoFatto As Boolean)

    Dim nrFile As Integer
    Dim riga As String

    On Error GoTo Gesterrore
    
    nrFile = FreeFile
    Open PathFileImportResult For Input As #nrFile
    Line Input #nrFile, riga
    
    Select Case riga
        Case "FAIL", ""
            'Recupero file fallito: devo rifarlo
            RecuperoFatto = False
            
        Case "OK"
            'Recupero file concluso con successo
            RecuperoFatto = True
        
        Case "TODO"
            'Recupero file in corso/da fare
            RecuperoFatto = RecuperoFatto
    End Select
    
    Close nrFile
    
    Exit Sub
Gesterrore:
    Call WindasLog("RecuperoDatiADAM5560ReadFile: " & Error(Err()), 1, "OPC")

End Sub

'Federica settembre 2017 - Scrittura del risultato del recupero dati su file
Public Sub RecuperoDatiADAM5560Result(ByVal TestoDaScrivere As String)
    
    Dim nrFile As Integer

    On Error GoTo Gesterrore

    nrFile = FreeFile()
    Open PathFileImportResult For Output As #nrFile
    Print #nrFile, TestoDaScrivere
    Close #nrFile
    
    Exit Sub
    
Gesterrore:
    Call WindasLog("RecuperoDatiADAM5560Result: " & Error(Err), 1, "OPC")

End Sub

'Alby Agosto 2017 - Scrittura su database dei dati recuperati
Public Function RecuperoDatiADAM5560ScriviDB(Dati) As Boolean

    Dim iDigitale As Integer
    Dim ValoreAcquisito
    Const OffSetDI = 18
    Dim iParametro As Integer

    On Error GoTo Gesterrore
    
    TimeStamp = Mid(Dati(0), 7, 2) + "/" + Mid(Dati(0), 5, 2) + "/" + Mid(Dati(0), 1, 4) + " " + Mid(Dati(0), 10, 2) + ":" + Mid(Dati(0), 12, 2) + ":" + Mid(Dati(0), 14, 2)
    
    'Federica settembre 2017 - Chiamata per i parametri stimati
    Call AcquisisceMisureStimate(ElencoMisureStimate)
        
    For iParametro = 0 To gnNroParametriStrumenti
        With ParametriStrumenti(iParametro)
            If .Acquisizione Then
                If (.FSE <> .ISE) Then
                    If .NroMorsetto >= 0 Then
                        
                        ValoreAcquisito = Dati(.NroMorsetto + 1)
                        
                        Status(0, iParametro) = "VAL"
                        Status(1, iParametro) = "VAL"
                        ValIst(0, iParametro) = .ISI + (ValoreAcquisito - .ISE) * (.FSI - .ISI) / (.FSE - .ISE)
                        
                        'Alby Agosto 2017
                        If .FattoreConversione > 0 Then ValIst(0, iParametro) = ValIst(0, iParametro) * .FattoreConversione
                        
                        ValIst(1, iParametro) = ValIst(0, iParametro)
                        
                        For iDigitale = 0 To nroDigitali
                            Valore_DI(iDigitale) = Dati(OffSetDI + iDigitale)
                            If Not StatoLogico_DI(iDigitale) Then
                                Valore_DI(iDigitale) = Abs(Valore_DI(iDigitale) - 1)
                            End If
                            
                            If Trim(CodiceParametro_DI(iDigitale)) <> "" Then
                                If InStr(.Invalida, Trim(CodiceParametro_DI(iDigitale))) > 0 Then
                                    If Valore_DI(iDigitale) = 1 Then
                                        Status(0, iParametro) = "ERR"
                                        Status(1, iParametro) = "ERR"
                                    End If
                                End If
                            End If
                        Next iDigitale
                    End If
                End If
            End If
        End With
    Next iParametro
    
    'Federica settembre 2017 - Calcolo dello Stato Impianto
    Call StatoImpianto
    ValIst(1, IngressoStatoImpianto) = ValIst(0, IngressoStatoImpianto)
    Status(1, IngressoStatoImpianto) = Status(0, IngressoStatoImpianto)
    
    'Federica settembre 2017 - Normalizzazione
    For iParametro = 0 To gnNroParametriStrumenti
        Call ElaborazioniDiLegge(iParametro, ValIst(1, iParametro), Status(1, iParametro))
    Next iParametro
    
    Call SalvaDatiElementariDB(True)
    
    RecuperoDatiADAM5560ScriviDB = True
    
    Exit Function

Gesterrore:

    Call WindasLog("RecuperoDatiADAM5560ScriviDB " + Error(Err), 1, OPC)

End Function


