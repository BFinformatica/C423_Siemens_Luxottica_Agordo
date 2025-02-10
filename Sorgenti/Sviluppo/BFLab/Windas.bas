Attribute VB_Name = "Windas"
Option Explicit

Global MancaComunicazionePC As Boolean

Sub WindasLog(Evento As String, grave, Form)

    Dim ll As Integer
    Dim rsConfig As Object
    Dim Tipo(1) As String
    
    On Error GoTo GestErrore
    
    Tipo(0) = "\logBFLab\"
    Tipo(1) = "\erroriBFLab\"
    
    Debug.Print Evento
    
    'Alby Febbraio 2012 log a video e su files giornalieri
    For ll = 0 To 1
        If Dir(App.Path & Tipo(ll), vbDirectory) = "" Then
            MkDir App.Path & Tipo(ll)
        End If
    Next ll
    
    ll = FreeFile
    Open App.Path & Tipo(grave) + Format(Now, "dd-mmmyy") + ".txt" For Append As #ll
    Print #ll, Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + Evento
    Close (ll)
    Form.TextLog = Format(Now, "dd/mm/yyyy hh.nn.ss") + " " + Evento + Chr(13) + Chr(10) + Left(Form.TextLog, 80000)
    If grave = 1 Then
        Form.TextLog.ForeColor = WHITE
        Form.TextLog.BackColor = RED
    Else
        Form.TextLog.ForeColor = WHITE
        Form.TextLog.BackColor = DARK_GREEN
    End If
    
    Exit Sub
    
GestErrore:
    Debug.Print Error(Err)
    
End Sub

Sub Ritardo(Secondi)
    
    'daniele agosto 2013 bolgiano: aggiungo nuova sub
    Dim jj
    Dim OldSecondi
    
    On Error GoTo GestErrore
    
    OldSecondi = second(Now)
    'Call GestioneErrore(err,"Ritardo1 Errore", "Tutto Regolare")
    
    For jj = 1 To Secondi
        Do
            DoEvents
            If second(Now) <> OldSecondi Then
                OldSecondi = second(Now)
                Exit Do
            End If
        Loop
    Next
    Exit Sub

GestErrore:
    Call WindasLog("Ritardo ", 1, OPC)
    
End Sub

Function LeggiINIfile(NomeFile, TagText)

    Dim ll As Integer
    Dim Linea As String
    Dim P As Integer
    
    On Error GoTo GestErrore
    
    'Alby Ottobre 2015
    LeggiINIfile = ""
    ll = FreeFile
    Open App.Path + "\" + NomeFile For Input As #ll

    Do While Not EOF(ll)
        Line Input #ll, Linea
        If InStr(UCase(Linea), UCase(TagText)) > 0 Then
            P = InStr(Linea, "=")
            If P > 0 Then LeggiINIfile = Mid(Linea, P + 1)
        End If
    Loop
    Close (ll)
 
    Exit Function
    
GestErrore:
    Call WindasLog("LeggiINIfile " + Error(Err) + " " + NomeFile + " " + TagText, 1, OPC)
    Close (ll)


End Function
