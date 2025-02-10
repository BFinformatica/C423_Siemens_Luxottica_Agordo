Attribute VB_Name = "GestioneSonoro"
Option Explicit

'luca luglio 2017
Private Declare Function playa Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Global PrimaVoltaAssolutoDigitali As Boolean
Private Const SND_FLAG = &H1 Or &H2

'luca luglio 2017 gestione sonoro allarmi
Sub GestioneAllarmiSonori()

    Dim i As Integer
    Dim SonoroAttivo As Boolean
    Dim TacitazioneSonoro As Integer
    
    On Error GoTo GestErrore
    
    SonoroAttivo = False
    
    TacitazioneSonoro = CInt(LeggiTag("DisattivazioneAllarmiSonori"))
    
    For i = 0 To nroDigitali
        If Sonoro_DI(i) = 1 And Not PrimaVoltaAssolutoDigitali Then
            If Valore_DI(i) = 1 Then
                SonoroAttivo = True
                Exit For
            End If
        End If
    Next i
    
    If Not PrimaVoltaAssolutoDigitali Then
        If SonoroAttivo And TacitazioneSonoro = 0 Then
            Call GestioneAllarmiSonoriAttiva(Generiche(iFileSonoro).Testo)
        End If
    End If
    
    PrimaVoltaAssolutoDigitali = False
Exit Sub
    
GestErrore:
    Call WindasLog("GestioneAllarmiSonori " + Error(Err), 1, OPC)
End Sub

'luca luglio 2017
Private Sub GestioneAllarmiSonoriAttiva(PercorsoFileAudio As String)
    
    On Error GoTo GestErrore
    
    If Dir$(PercorsoFileAudio) <> "" Then playa PercorsoFileAudio, SND_FLAG
    
    Exit Sub
    
GestErrore:
    Call WindasLog("GestioneAllarmiSonoriAttiva " + Error(Err), 1, OPC)
End Sub
