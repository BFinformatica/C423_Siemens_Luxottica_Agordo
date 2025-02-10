VERSION 5.00
Begin VB.Form OPC 
   Caption         =   "BFwinCC"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   Icon            =   "OPCfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextLog 
      Height          =   3615
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   13215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12480
      Top             =   120
   End
   Begin VB.Label TimeDate 
      Caption         =   "TimeDate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "OPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mcp As Object
Dim VersioneTag As String 'Alby Dicembre 2015

'Alby Maggio 2016 OPC
Dim Server As OpcClientX.OPCServer
Dim WithEvents group As OpcClientX.OPCGroup
Attribute group.VB_VarHelpID = -1
Dim Item(1000) As OpcClientX.OPCItem
Dim TagName(1000) As String
Dim MaxTag As Integer

Sub InizializzaWinCC()

    On Error GoTo GestErrore

    Call WindasLog("Inizilizzo connessione a WinCC", 0, OPC)
    Set mcp = Nothing
    Set mcp = CreateObject("WinCC-Runtime-Project")
        
    Exit Sub
    
GestErrore:
    Call WindasLog("InizializzaWinCC " + Error(Err), 1, OPC)

End Sub

Function ControllaTag(NomeTag) As Integer

    Dim iIdx As Integer

    'Alby Maggio 2016 OPC
    On Error GoTo GestErrore

    ControllaTag = -1
    For iIdx = 0 To MaxTag
        If TagName(iIdx) = NomeTag Then
            ControllaTag = iIdx
            Exit Function
        End If
    Next iIdx

    Exit Function
    
GestErrore:
    Call WindasLog("ControllaTag " + Error(Err), 1, OPC)

End Function

Function LeggiTagWinCC(NomeTag)

    Dim ErroreConnessione

    On Error GoTo GestErrore

Riprova:
    
    'Alby Maggio 2016 OPC
    'luca giugno 2017
    #If versione = 1 Then
        If VersioneTag = "OPC" Then
            LeggiTagWinCC = LeggiTagOPC(NomeTag)
            Exit Function
        Else
            LeggiTagWinCC = mcp.getValue(NomeTag)
        End If
    #Else
        LeggiTagWinCC = LeggiTagOPC(NomeTag)
        Exit Function
    #End If
    
    'Alby Febbraio 2014 se manca WinCC
    If IsEmpty(LeggiTagWinCC) Then
        'End

ErroreWinCC:
        
        'Alby Aprile 2014
        'Debug.Print NomeTag
        
        ErroreConnessione = 1
        Call WindasLog("Attenzione WinCC non operativo! (Err. su tag " & NomeTag, 1, OPC)
        Call InizializzaWinCC
        Call Ritardo(10)
        Call InizializzaSistema
        GoTo Riprova
    Else
        If ErroreConnessione = 1 Then
            Call WindasLog("Connessione con WinCC ripristinata", 0, OPC)
        End If
    End If
    
    Exit Function
    
GestErrore:
    If Err = 462 Then
        Call WindasLog("Attenzione WinCC non operativo! ", 1, OPC)
        GoTo ErroreWinCC
    Else
        Call WindasLog("LeggiTagWinCC " + Error(Err), 1, OPC)
    End If
    
End Function

Sub ScriviTagWinCC(NomeTag, Variabile)

    Dim StrNomeTag As String

    On Error GoTo GestErrore
      
    'Alby Maggio 2016 OPC
    'luca giugno 2017
    #If versione = 1 Then
        If VersioneTag = "OPC" Then
            Call ScriviTagOPC(NomeTag, Variabile)
            Exit Sub
        Else
            mcp.setvalue NomeTag, Variabile
        End If
    #Else
        Call ScriviTagOPC(NomeTag, Variabile)
        Exit Sub
    #End If
    
    Exit Sub
    
GestErrore:
    If Err = 462 Then
        Call WindasLog("WinCC non operativo! ", 1, OPC)
    Else
        Call WindasLog("GestioneTag " + Error(Err), 1, OPC)
    End If

End Sub

'luca gennaio 2018
Sub ChiudiOPC()

'luca febbraio 2018
Dim i As Integer

On Error GoTo GestErrore
    
    If Not Server Is Nothing Then
        
        'luca febbraio 2018
        For i = 0 To UBound(TagName)
            TagName(i) = ""
            Set Item(i) = Nothing
        Next i
        MaxTag = 0
        
        Server.OPCGroups.RemoveAll
        Set group = Nothing
        
        Server.Disconnect
        Set Server = Nothing
    End If
    
Exit Sub
    
GestErrore:
    Call WindasLog("ChiudiOPC " + Error(Err), 1, OPC)
    Set group = Nothing
    Set Server = Nothing
End Sub

Private Sub Form_Load()
    
    'VERSIONE 1 = WINCC SiCEMS
    'VERSIONE 2 = WINCC ADVANCED WINDAS
    
    If App.PrevInstance = True Then End
    
    Client = IsClient
    
    Call WindasLog("Avviato modulo BFLab", 0, OPC)

    'Alby Giugno 2016 OPC
    VersioneTag = "OPC"
    DaLeggereRegistriPLC = True
    
    #If versione = 1 Then
        If VersioneTag = "OPC" Then
            Call InizializzaOPC
        Else
            Call InizializzaWinCC
        End If
    #ElseIf versione = 3 Then
        'Alby Marzo 2018 Versione WinDas.net Classic con GUI in VisualBasic.net
        Call InizializzaComunicator
        
    #Else
        'luca giugno 2017 nella versione con wincc advanced non posso utilizzare l'inizializzazione con l'oggetto di WinCC
        Call InizializzaOPC
    #End If
    
    

    
    Call InizializzaSistema 'Federica gennaio 2018
    
    'Alby Agosto 2017
    Me.Visible = False
    Me.Timer1.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.WindowState = MINIMIZED
    
    Cancel = True

End Sub


Private Sub Timer1_Timer()

    TimeDate.Caption = Format(Now, "dd/mm/yyyy hh.nn.ss")
    Call Acquisisce
    
End Sub

Sub ScriviTagOPC(NomeTag, Variabile)

    Dim StrNomeTag As String
    Dim IndiceTag As Integer

    On Error GoTo GestErrore
       
    'Alby Maggio 2016 OPC
    IndiceTag = ControllaTag(NomeTag)
    If IndiceTag >= 0 Then
        'luca 08/11/2016 controllo su variabile blank
        If Not IsNull(Variabile) And Not IsEmpty(Variabile) Then
            Item(IndiceTag).Write Variabile
        End If
        'Debug.Print Format(Now, "hh.nn.ss") + " scrivo variabile"
        Exit Sub
    Else
        Debug.Print Format(Now, "hh.nn.ss") + " aggiungo tag"
        TagName(MaxTag) = NomeTag
        Set Item(MaxTag) = group.OPCItems.AddItem(NomeTag, 0)
    End If
    
    'luca 08/11/2016 controllo su variabile blank
    If Not IsNull(Variabile) And Not IsEmpty(Variabile) Then
        Item(MaxTag).Write Variabile
    End If
    
    MaxTag = MaxTag + 1
    If MaxTag > 999 Then
        Call WindasLog("Attenzione raggiunto numero max di tag OPC", 1, OPC)
    End If
    
    Exit Sub
    
GestErrore:

    'Resume Next
    
    Call WindasLog("ScriviTagOPC: <" & NomeTag & "> " & Error(Err), 1, OPC)
    Resume Next
    Call InizializzaOPC
    
End Sub

Sub InizializzaOPC()

    Dim iIdx As Integer
    Dim NumeroTentativi As Integer  'Luca gennaio 2018 - Gestione ripetizione inizializzazione

    On Error GoTo GestErrore
    
    'Luca gennaio 2018
    Call ChiudiOPC

Reinizializza:
    Set Server = New OpcClientX.OPCServer
   
    ' connect to server
    'luca giugno 2017
    #If versione = 1 Then
        Server.Connect "OPCServer.WinCC.1"
    #Else
        Server.Connect "OPC.SimaticHMI.CoRtHmiRTm.1"
    #End If
    
    If Server.ServerName = "" Then Exit Sub

    Set group = Server.OPCGroups.Add("Group1")

    Exit Sub
    
GestErrore:
    'luca gennaio 2018 - all'avvio il BFLab spesso non riesce ad attaccarsi e
    'si rischia di comprometterne il funzionamento fin tanto che non si riavvia
    '(Scrivitag/leggitag vanno in errore), così solo all'avvio (form_load) tenta per 5 volte di attaccarsi
    NumeroTentativi = NumeroTentativi + 1
    If NumeroTentativi < 5 Then
        Call Ritardo(1)
        GoTo Reinizializza
    Else
        'Federica febbraio 2018 - Chiudo BFLab per evitare ulteriori errori.
        End
    End If

End Sub

'luca 08/11/2016
Function IsClient() As Boolean

    Dim nFile As Integer
    Dim riga As String
    
    On Error GoTo GestErrore
    
    'Alby Luglio 2016
    'se presente file ini
    nFile = FreeFile
    Open App.Path & "\Client.ini" For Input As #nFile
    Line Input #nFile, riga
    'Alby Gennaio 2016
    IsClient = CBool(Trim(riga))
    Close #nFile
    
    Exit Function
    
GestErrore:
    Call WindasLog("IsClient " + Error(Err), 1, OPC)

End Function

Function LeggiTagOPC(NomeTag)

    Dim StrNomeTag As String
    Dim IndiceTag As Integer
    Dim TagAggiornata As Date
    
    On Error GoTo GestErrore

    'Alby Maggio 2016 OPC
    IndiceTag = ControllaTag(NomeTag)
    If IndiceTag >= 0 Then
        
        Item(IndiceTag).Read 1
        LeggiTagOPC = Item(IndiceTag).Value
        Exit Function
    Else
        Debug.Print Format(Now, "hh.nn.ss") + " aggiungo tag"
        TagName(MaxTag) = NomeTag
        Set Item(MaxTag) = group.OPCItems.AddItem(NomeTag, 0)
    End If
    
    Item(MaxTag).Read 1
     
    LeggiTagOPC = Item(MaxTag).Value
    
    MaxTag = MaxTag + 1
    If MaxTag > 999 Then
        Call WindasLog("Attenzione raggiunto numero max di tag OPC", 1, OPC)
    End If
    
    Exit Function
    
GestErrore:
    
    Call WindasLog("LeggiTagOPC: <" & NomeTag & "> " & Error(Err), 1, OPC)
    Call InizializzaOPC

End Function

