VERSION 5.00
Begin VB.Form OPC 
   Caption         =   "BFwinCC"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
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
      Interval        =   50
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

'Alby Maggio 2016 OPC
Dim Server As OpcClientX.OPCServer
Dim WithEvents group As OpcClientX.OPCGroup
Attribute group.VB_VarHelpID = -1
Dim Item(1000) As OpcClientX.OPCItem
Dim TagName(1000) As String
Dim MaxTag As Integer

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
    Call WindasLog("ControllaTag " + Error(Err), 1)

End Function

Sub ScriviTagOPC(NomeTag, Variabile)

    Dim StrNomeTag As String
    Dim IndiceTag As Integer

    On Error GoTo GestErrore
    
    If OPCinErrore Then Exit Sub
       
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
    
    Item(MaxTag).Write Variabile
    MaxTag = MaxTag + 1
    If MaxTag > 999 Then
        Call WindasLog("Attenzione raggiunto numero max di tag OPC", 1)
    End If
    
    Exit Sub
    
GestErrore:

    Call WindasLog("ScriviTagOPC :" & NomeTag & " - " & Error(Err), 1)
    Resume Next
    Call InizializzaOPC
    
End Sub

Sub InizializzaOPC()

    Dim iIdx As Integer

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
    
    OPCinErrore = False     'Federica novembre 2017

    Set group = Server.OPCGroups.Add("Group1")

    Exit Sub
    
GestErrore:
    'Federica novembre 2017 - Gestione errore OPC
    Select Case Err
        Case 429
            Call WindasLog("InizializzaOPC: OPC in errore o non avviato. " + Error(Err), 1)
            OPCinErrore = True
        Case Else
            Call WindasLog("InizializzaOPC " + Error(Err), 1)
    End Select

End Sub

Function LeggiTagOPC(NomeTag)

    Dim StrNomeTag As String
    Dim IndiceTag As Integer
    
    On Error GoTo GestErrore
    
    If OPCinErrore Then Exit Function

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
        Call WindasLog("Attenzione raggiunto numero max di tag OPC", 1)
    End If
    
    Exit Function
    
GestErrore:
    
    Call WindasLog("LeggiTagOPC :" & NomeTag & " - " & Error(Err), 1)
    Call InizializzaOPC

End Function

'luca luglio 2017
Sub ChiudiOPC()

    Dim i As Integer    'Luca febbraio 2018

    On Error GoTo GestErrore
    
    If Not Server Is Nothing Then
    
        'Luca febbraio 2018
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
    Call WindasLog("ChiudiOPC " + Error(Err), 1)
    Set group = Nothing
    Set Server = Nothing

End Sub

