Attribute VB_Name = "Dizionario"
Public LocalizationDictionary As Dictionary

Function Loc(StringaLoc As String) As String

    On Error GoTo GestErrore
    Loc = ""
    If StringaLoc = "" Then Exit Function
    
    LocalizationDictionary.CompareMode = BinaryCompare
    
    'Se esiste la chiave
    If LocalizationDictionary.Exists(StringaLoc) Then
        Loc = LocalizationDictionary.Item(StringaLoc)
    Else
        Loc = "---"
    End If

    Exit Function

GestErrore:
    Call CReport.windasLog("Loc: " & Error(Err))
    Loc = "---"

End Function

Sub CaricaFileTraduzione()

    Dim ll As Integer
    Dim Chiave As String
    Dim Traduzione As String
    Dim TempStr() As String
    Dim strBData As String
    Dim Linee() As String
    Dim conta As Integer
    
    On Error GoTo GestErrore
    
    Set LocalizationDictionary = New Dictionary
    
    'Se c'è il file di traduzione
    If Dir(App.Path & "\LocTable.ini") <> "" Then
        'Apertura del file
        ll = FreeFile
        'Open App.Path & "\LocTable.ini" For Input As #ll
        Open App.Path & "\LocTable.ini" For Binary As #ll
        strBData = InputB(LOF(ll), #ll)
        Close #ll
        Linee = Split(strBData, vbCrLf)
        For conta = 0 To UBound(Linee)
            TempStr = Split(Linee(conta), Chr(9))
            If conta = 0 Then
                'Elimino i primi 2 caratteri
                TempStr(0) = Mid$(TempStr(0), 2)
            End If
            Chiave = TempStr(0)
            If UBound(TempStr) > 0 Then
                Traduzione = TempStr(1)
            Else
                Traduzione = ""
            End If
            'Caricamento nel dictionary delle traduzioni (solo se valide entrambe)
            If Chiave <> "" And Traduzione <> "" Then
                Call LocalizationDictionary.Add(Chiave, Traduzione)
            End If
        Next conta
    Else
        Call CReport.windasLog("CaricaFileTraduzione: Nessun file traduzione presente")
        Exit Sub
    End If
    
    Exit Sub
    
GestErrore:
    Call CReport.windasLog("CaricaFileTraduzione: " & Error(Err))
    Resume Next
    Close (ll)

End Sub
