VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Elaborazione dati SiCEMS"
   ClientHeight    =   5484
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   8076
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   8076
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Dati Minuto (solo DB)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   19
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dati Orari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   18
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dati Semiorari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dati 10 Minuti CO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   16
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox cbo_giorno_fin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cbo_mese_fin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cbo_anno_fin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rielabora dati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   3375
   End
   Begin VB.ComboBox cbo_anno_in 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbo_mese_in 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbo_giorno_in 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "al"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Giorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Mese"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Anno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Dal"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Anno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Mese"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Giorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)

    Dim nFile As Integer
    Dim nn As Integer
    
    On Local Error GoTo GestErrore
    
    
    DatiDaElaborare(Index) = Check1(Index).Value
    
    nFile = FreeFile
    Open App.Path & "\BFData.ini" For Output As #nFile
    For nn = 0 To 3
        Print #nFile, Trim(Str(DatiDaElaborare(nn)))
    Next nn
    Close #nFile
    
    Exit Sub
    
GestErrore:

    Debug.Print Now & " Check1_Click: " & Error(Err)
    Close #nFile
    
End Sub

Private Sub Command1_Click()

    Dim DataIn As Date
    Dim DataFin As Date
    Dim giorno As String
    Dim mese As String
    Dim anno As String
    Dim Passato As Boolean
    Dim nn As Integer
    
    
    Passato = False
    
    For nn = 0 To 3
        If DatiDaElaborare(nn) Then
            Passato = True
        End If
    Next nn
    
    If Passato Then
    
    
        If MsgBox("Avvio la rielaborazione?", vbOKCancel Or vbQuestion) = vbOK Then
        
            '**** Lettura configurazione
            Form1.Label1.Caption = "Lettura configurazione..."
            Form1.Refresh
            Call LeggiConfigurazione7
            
            '**** Determina periodo di elaborazione
            giorno = cbo_giorno_in.Text
            mese = cbo_mese_in.Text
            anno = cbo_anno_in.Text
            DataIn = CDate(giorno & "/" & mese & "/" & anno)
            
            giorno = cbo_giorno_fin.Text
            mese = cbo_mese_fin.Text
            anno = cbo_anno_fin.Text
            DataFin = CDate(giorno & "/" & mese & "/" & anno)
          
            For nn = 0 To 3
            
                If DatiDaElaborare(nn) Then
                    
                    Select Case nn
                        Case 0
                            Tabella = "WDS_10MINCO"
                            StrLabel = "Elaborazioni dati 10 minuti: "
                            'Label1.Caption = "Elaborazioni dati 10 minuti..."
                            
                        Case 1
                            Tabella = "WDS_HALF"
                            StrLabel = "Elaborazioni dati semiorari: "
                            
                        Case 2
                            Tabella = "WDS_ELAB"
                            StrLabel = "Elaborazioni dati orari: "
                              
                        Case 3
                            Tabella = "WDS_AUTO"
                            StrLabel = "Elaborazioni dati minuto: "
                    End Select
        
                    
                    ElabDate = DataIn
                    Do
                      Call Elabora(ElabDate)
                      ElabDate = DateAdd("d", 1, ElabDate)
                    Loop While ElabDate <= DataFin
                
                End If
                
            Next nn
            
            Label1.Caption = "Rialoborazione dati completata."
            Form1.Refresh
          
        End If
        
    Else
        MsgBox ("Attenzione, nessuna elaborazione selezionata!")
        Label1.Caption = ""
    End If
    

End Sub

Private Sub Form_Load()

    Dim yy As Integer
    Dim dd As Integer
    Dim mm As Integer
    'Dim auto As Boolean
    Dim rs As Object
    Dim nn As Integer
    Dim nFile As Integer
    Dim riga As String
    
    On Error GoTo GestErrore
    
    'Alby Giugno 2016
    VersioneTag = "OPC"
    Ruolo = RicavoRuolo
    Client = IsClient
    
    Call WindasLog("Avviato BFdata", 0)
    
    'Parametri cdi connessione al db
    Call GetConnectionParam
    NewDataObj rs
    
    rs.SelectionFast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='DIR_LAV'"
    gsDirLavoro = rs.getValue("gt_Value")

    '***** Numero Linea *****
    'Federica gennaio 2018 - Lettura numero linea da DB
    rs.SelectionFast "select * from wds_gentab where gt_type='stations' and gt_code='" & StationCode & "'"
    If Not rs.IsEOF Then
        NumeroLineaBFData = rs.getValue("gt_order")
    Else
        'Se manca il dato esco
        MsgBox "Manca numero linea per la stazione " & StationCode & " nel DataBase. Esecuzione terminata!", vbCritical
        End
    End If
    
    'Alby Dicembre 2015
    'Path file unico arpa
    rs.SelectionFast "select gt_value from wds_gentab where gt_type = 'opparm' and gt_code ='4343_FLD'"
    PathARPA_FileUnico = rs.getValue("gt_value")

    Set rs = Nothing
    
    '***** lettura file BFdata.ini *****
    If Dir(App.Path & "\BFData.ini") <> "" Then
    
        nFile = FreeFile
        Open App.Path & "\BFData.ini" For Input As #nFile
        For nn = 0 To 3
            Line Input #nFile, riga
            'Alby Gennaio 2016
            If LCase(riga) = "false" Or LCase(riga) = "falso" Then
                DatiDaElaborare(nn) = False
            Else
                DatiDaElaborare(nn) = True
            End If
        Next nn
        Close #nFile
    
    Else
    
        For nn = 0 To 3
            DatiDaElaborare(nn) = False
        Next nn
    
    End If
    
    For nn = 0 To 3
        If DatiDaElaborare(nn) Then
            Check1(nn).Value = 1
        Else
            Check1(nn).Value = 0
        End If
    Next nn
    
    '***** BFData lanciato da BFLab *****
    If InStr(UCase(Command), "AUTO") > 0 Then
    
        Call LeggiConfigurazione7
        'daniele luglio 2013 bolgiano: data su elabdate
        ElabDate = Now
        'ElabDate = Date
        
        For nn = 0 To 2
        
             If InStr((Command), Trim(Str(nn))) > 0 Then
            
                Select Case nn
                    Case 0
                        Tabella = "WDS_10MINCO"
                        
                    Case 1
                        Tabella = "WDS_HALF"
                        
                    Case 2
                        Tabella = "WDS_ELAB"
                End Select
        
                'se mezzanotte rielaboro anche giorno precedente
                'daniele luglio 2013 bolgiano
                'If Hour(Now) = 0 Then
                If hour(ElabDate) = 0 Then
                    If nn = 0 Then
                        '***** medie 10 minuti CO *****
                        'If Minute(Now) < 10 Then
                        If minute(ElabDate) < 10 Then
                            'Call Elabora(DateAdd("d", -1, Now))
                            'luca aprile 2017
                            'Call Elabora(DateAdd("d", -1, ElabDate))
                            ElabDate = DateAdd("d", -1, Now)
                            Call Elabora(ElabDate)
                        Else
                            '***** elaboro giornata in corso *****
                            'Call Elabora(Now)
                            Call Elabora(ElabDate)
                        End If

                    ElseIf nn = 1 Then
                        '***** medie semiorarie *****
                        'If Minute(Now) < 30 Then
                        If minute(ElabDate) < 30 Then
                            'Call Elabora(DateAdd("d", -1, Now))
                            'luca aprile 2017
                            'Call Elabora(DateAdd("d", -1, ElabDate))
                            ElabDate = DateAdd("d", -1, Now)
                            Call Elabora(ElabDate)
                        Else
                            '***** elaboro giornata in corso *****
                            'Call Elabora(Now)
                            Call Elabora(ElabDate)
                        End If

                    ElseIf nn = 2 Then
                        'Alby Dicembre 2015
                        ElabDate = DateAdd("d", -1, Now)
                        Call Elabora(ElabDate)
                        

                    End If
                Else
                    '***** elaboro giornata in corso *****
                    'Call Elabora(Now)
                    Call Elabora(ElabDate)
                End If
                
            End If
            
        Next nn
        
        Unload Me
        
        End
        
    Else
        
        '***** lancio normale con interfaccia *****
        For nn = 1 To 31
            cbo_giorno_in.AddItem Format(nn, "00")
            If nn <= 12 Then cbo_mese_in.AddItem Format(nn, "00")
            yy = 2006 + nn
            cbo_anno_in.AddItem Trim$(Str$(yy))
        Next nn
        For nn = 1 To 31
            cbo_giorno_fin.AddItem Format(nn, "00")
            If nn <= 12 Then cbo_mese_fin.AddItem Format(nn, "00")
            yy = 2006 + nn
            cbo_anno_fin.AddItem Trim$(Str$(yy))
        Next nn

        'luca 15/09/2016
'        dd = Val(Left$(Now, 2))
'        mm = Val(Mid(Now, 4, 2))
'        yy = Val(Mid$(Now, 7, 4))

        dd = day(Now)
        mm = month(Now)
        yy = year(Now)
    
        cbo_giorno_in.ListIndex = dd - 1
        cbo_mese_in.ListIndex = mm - 1
        cbo_anno_in.ListIndex = yy - 2007
        
        cbo_giorno_fin.ListIndex = dd - 1
        cbo_mese_fin.ListIndex = mm - 1
        cbo_anno_fin.ListIndex = yy - 2007

   End If

   Exit Sub
    
GestErrore:
    Debug.Print Now & " Form_Load: " & Error(Err)
    Call WindasLog("BFData Form_Load " + Error(Err), 1)
    Resume Next

fine:

    
End Sub
