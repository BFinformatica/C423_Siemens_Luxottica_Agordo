VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurazione mappatura I/O"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbLinea 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   2880
      List            =   "Form2.frx":0016
      TabIndex        =   12
      Text            =   "Linea"
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton cmd_esci 
      Caption         =   "&Esci"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   8400
      Width           =   1275
   End
   Begin VB.CommandButton cmd_salva 
      Caption         =   "&Salva"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7980
      TabIndex        =   9
      Top             =   8400
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14314
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingressi"
      TabPicture(0)   =   "Form2.frx":0050
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GridEX_AI"
      Tab(0).Control(1)=   "GridEX_DI"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "Label2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Uscite"
      TabPicture(1)   =   "Form2.frx":006C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "GridEX_DO"
      Tab(1).Control(3)=   "GridEX_AO"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Diagnostica"
      TabPicture(2)   =   "Form2.frx":0088
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "WatchDog su BFComunicator"
         Height          =   1815
         Left            =   360
         TabIndex        =   20
         Top             =   3240
         Width           =   4575
         Begin VB.CheckBox chk_watchdog 
            Caption         =   "Salva WatchDog su BFComunicator"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   3255
         End
         Begin VB.PictureBox pic_watchdog 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   4215
            TabIndex        =   21
            Top             =   960
            Visible         =   0   'False
            Width           =   4215
            Begin VB.TextBox txt_tag_wd 
               Height          =   285
               Left            =   1200
               TabIndex        =   22
               Text            =   "WATCHDOG"
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Tag utilizzata"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   135
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "BFComunicator"
         Height          =   735
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   4575
         Begin VB.CheckBox chk_bfcomunicator 
            Caption         =   "Aggiorna letture in BFComunicator"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Value           =   1  'Checked
            Width           =   3255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Allarmi di comunicazione"
         Height          =   1695
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   4575
         Begin VB.CheckBox chk_allarmi 
            Caption         =   "Salva allarmi di comunicazione"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.PictureBox pic_alarm 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4275
            TabIndex        =   14
            Top             =   840
            Width           =   4275
            Begin VB.TextBox txt_morsetto_allarme 
               Height          =   285
               Left            =   1200
               TabIndex        =   15
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label19 
               Caption         =   "Indice iniziale"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   255
               Width           =   1095
            End
         End
      End
      Begin GridEX20.GridEX GridEX_AI 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   1
         Top             =   960
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   5318
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Form2.frx":00A4
         Column(2)       =   "Form2.frx":016C
         FormatStylesCount=   6
         FormatStyle(1)  =   "Form2.frx":0210
         FormatStyle(2)  =   "Form2.frx":0348
         FormatStyle(3)  =   "Form2.frx":03F8
         FormatStyle(4)  =   "Form2.frx":04AC
         FormatStyle(5)  =   "Form2.frx":0584
         FormatStyle(6)  =   "Form2.frx":063C
         ImageCount      =   0
         PrinterProperties=   "Form2.frx":071C
      End
      Begin GridEX20.GridEX GridEX_DI 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   2
         Top             =   4920
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   5318
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Form2.frx":08F4
         Column(2)       =   "Form2.frx":09BC
         FormatStylesCount=   6
         FormatStyle(1)  =   "Form2.frx":0A60
         FormatStyle(2)  =   "Form2.frx":0B98
         FormatStyle(3)  =   "Form2.frx":0C48
         FormatStyle(4)  =   "Form2.frx":0CFC
         FormatStyle(5)  =   "Form2.frx":0DD4
         FormatStyle(6)  =   "Form2.frx":0E8C
         ImageCount      =   0
         PrinterProperties=   "Form2.frx":0F6C
      End
      Begin GridEX20.GridEX GridEX_AO 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   5
         Top             =   960
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   5318
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Form2.frx":1144
         Column(2)       =   "Form2.frx":120C
         FormatStylesCount=   6
         FormatStyle(1)  =   "Form2.frx":12B0
         FormatStyle(2)  =   "Form2.frx":13E8
         FormatStyle(3)  =   "Form2.frx":1498
         FormatStyle(4)  =   "Form2.frx":154C
         FormatStyle(5)  =   "Form2.frx":1624
         FormatStyle(6)  =   "Form2.frx":16DC
         ImageCount      =   0
         PrinterProperties=   "Form2.frx":17BC
      End
      Begin GridEX20.GridEX GridEX_DO 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   6
         Top             =   4920
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   5318
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "Form2.frx":1994
         Column(2)       =   "Form2.frx":1A5C
         FormatStylesCount=   6
         FormatStyle(1)  =   "Form2.frx":1B00
         FormatStyle(2)  =   "Form2.frx":1C38
         FormatStyle(3)  =   "Form2.frx":1CE8
         FormatStyle(4)  =   "Form2.frx":1D9C
         FormatStyle(5)  =   "Form2.frx":1E74
         FormatStyle(6)  =   "Form2.frx":1F2C
         ImageCount      =   0
         PrinterProperties=   "Form2.frx":200C
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mappatura uscite analogiche"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74880
         TabIndex        =   8
         Top             =   420
         Width           =   10125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Mappatura uscite digitali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74880
         TabIndex        =   7
         Top             =   4380
         Width           =   10125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mappatura ingressi analogici"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   10125
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Mappatura ingressi digitali"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74880
         TabIndex        =   3
         Top             =   4380
         Width           =   10125
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Linea di appartenenza:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   8400
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MappaturaAImodificata As Boolean
Dim MappaturaDImodificata As Boolean
Dim MappaturaAOmodificata As Boolean
Dim MappaturaDOmodificata As Boolean
Dim LineaModificata As Boolean
Dim ImpostazioniModificate As Boolean
Dim loading As Boolean

Private Sub chk_allarmi_Click()

  Call SettaModifiche
  
End Sub

Private Sub chk_bfcomunicator_Click()

    Call SettaModifiche
    
End Sub

Private Sub chk_watchdog_Click()

  pic_watchdog.Visible = chk_watchdog.Value
  Call SettaModifiche
  
End Sub

Private Sub cmbLinea_Change()

    LineaModificata = True
    cmd_salva.Enabled = True

End Sub

Private Sub cmbLinea_Click()

    LineaModificata = True
    cmd_salva.Enabled = True

End Sub

Private Sub cmd_esci_Click()
    
    Unload Me
    
End Sub

Private Sub cmd_salva_Click()

    Call SalvaModifiche

End Sub

Private Sub CaricaTabellaIngressiAnalogici()

    On Error GoTo GestErr
    
    '***** Mappatura ingressi analogici
    GridEX_AI.AllowAddNew = True
    GridEX_AI.Columns.Clear
    Call GridEX_AI.Columns.Add("Descrizione", jgexText, jgexEditTextBox, "PLC")
    Call GridEX_AI.Columns.Add("DB", jgexText, jgexEditTextBox, "DB")
    Call GridEX_AI.Columns.Add("Indirizzo base", jgexText, jgexEditTextBox, "ADDR")
    Call GridEX_AI.Columns.Add("Tipo variabile", jgexText, jgexEditDropDown, "TIPO")
    Call GridEX_AI.Columns.Add("N. ingressi", jgexText, jgexEditTextBox, "N_AI")
    Call GridEX_AI.Columns.Add("N. bytes ", jgexText, jgexEditTextBox, "NBYTES")
    Call GridEX_AI.Columns.Add("Range ingressi", jgexText, jgexEditTextBox, "RANGE_AI")
    
    GridEX_AI.Columns("TIPO").HasValueList = True
    GridEX_AI.Columns("TIPO").ValueList.Add 0, "INTEGER"
    GridEX_AI.Columns("TIPO").ValueList.Add 1, "FLOAT"
    
    GridEX_AI.Columns("NBYTES").Selectable = False
    GridEX_AI.Columns("RANGE_AI").Selectable = False
    
    CaricaMappaAI
    GridEX_AI.ItemCount = mRecordCount
    GridEX_AI.Rebind

    Exit Sub
    
GestErr:
    Debug.Print "CaricaTabellaIngressiAnalogici(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaTabellaIngressiAnalogici(): " & Err.Description)
    Resume Next

End Sub

Private Sub CaricaTabellaIngressiDigitali()

    On Error GoTo GestErr
    
    '***** Mappatura ingressi digitali
    GridEX_DI.AllowAddNew = True
    GridEX_DI.Columns.Clear
    Call GridEX_DI.Columns.Add("Descrizione", jgexText, jgexEditTextBox, "PLC")
    Call GridEX_DI.Columns.Add("DB", jgexText, jgexEditTextBox, "DB")
    Call GridEX_DI.Columns.Add("Indirizzo base", jgexText, jgexEditTextBox, "ADDR")
    Call GridEX_DI.Columns.Add("N. bytes ", jgexText, jgexEditTextBox, "NBYTES")
    Call GridEX_DI.Columns.Add("Tipo variabile", jgexText, jgexEditDropDown, "TIPO")
    Call GridEX_DI.Columns.Add("N. ingressi", jgexText, jgexEditTextBox, "N_DI")
    Call GridEX_DI.Columns.Add("Range ingressi", jgexText, jgexEditTextBox, "RANGE_DI")
    
    GridEX_DI.Columns("TIPO").Visible = False
    
    GridEX_DI.Columns("N_DI").Selectable = False
    GridEX_DI.Columns("RANGE_DI").Selectable = False
    
    CaricaMappaDI
    GridEX_DI.ItemCount = mRecordCountDI
    GridEX_DI.Rebind

    Exit Sub
    
GestErr:
    Debug.Print "CaricaTabellaIngressiDigitali(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaTabellaIngressiDigitali(): " & Err.Description)
    Resume Next

End Sub

Private Sub CaricaTabellaUsciteDigitali()

    On Error GoTo GestErr

    '***** Mappatura uscite digitali
    GridEX_DO.Columns.Clear
    GridEX_DO.AllowAddNew = True
    Call GridEX_DO.Columns.Add("Descrizione", jgexText, jgexEditTextBox, "PLC")
    Call GridEX_DO.Columns.Add("DB", jgexText, jgexEditTextBox, "DB")
    Call GridEX_DO.Columns.Add("Indirizzo base", jgexText, jgexEditTextBox, "ADDR")
    Call GridEX_DO.Columns.Add("N. bytes ", jgexText, jgexEditTextBox, "NBYTES")
    Call GridEX_DO.Columns.Add("Tipo variabile", jgexText, jgexEditDropDown, "TIPO")
    Call GridEX_DO.Columns.Add("N. uscite", jgexText, jgexEditTextBox, "N_DO")
    Call GridEX_DO.Columns.Add("Range uscite", jgexText, jgexEditTextBox, "RANGE_DO")
    
    GridEX_DO.Columns("TIPO").Visible = False
    
    GridEX_DO.Columns("N_DO").Selectable = False
    GridEX_DO.Columns("RANGE_DO").Selectable = False
    
    CaricaMappaDO
    GridEX_DO.ItemCount = mRecordCountDO
    GridEX_DO.Rebind
    
    Exit Sub
    
GestErr:
    Debug.Print "CaricaTabellaUsciteDigitali(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaTabellaUsciteDigitali(): " & Err.Description)
    Resume Next

End Sub

Private Sub CaricaTabellaUsciteAnalogiche()

    On Error GoTo GestErr

    '***** Mappatura uscite analogiche
    GridEX_AO.Columns.Clear
    GridEX_AO.AllowAddNew = True
    Call GridEX_AO.Columns.Add("Descrizione", jgexText, jgexEditTextBox, "PLC")
    Call GridEX_AO.Columns.Add("DB", jgexText, jgexEditTextBox, "DB")
    Call GridEX_AO.Columns.Add("Indirizzo base", jgexText, jgexEditTextBox, "ADDR")
    Call GridEX_AO.Columns.Add("Tipo variabile", jgexText, jgexEditDropDown, "TIPO")
    Call GridEX_AO.Columns.Add("N. uscite", jgexText, jgexEditTextBox, "N_AO")
    Call GridEX_AO.Columns.Add("N. bytes ", jgexText, jgexEditTextBox, "NBYTES")
    Call GridEX_AO.Columns.Add("Range uscite", jgexText, jgexEditTextBox, "RANGE_AO")
    
    GridEX_AO.Columns("TIPO").HasValueList = True
    GridEX_AO.Columns("TIPO").ValueList.Add 0, "INTEGER"
    GridEX_AO.Columns("TIPO").ValueList.Add 1, "FLOAT"
    
    GridEX_AO.Columns("NBYTES").Selectable = False
    GridEX_AO.Columns("RANGE_AO").Selectable = False
    
    CaricaMappaAO
    GridEX_AO.ItemCount = mRecordCountAO
    GridEX_AO.Rebind
    Exit Sub
    
GestErr:
    Debug.Print "CaricaTabellaUsciteAnalogiche: " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaTabellaUsciteAnalogiche: " & Err.Description)
    Resume Next

End Sub

Private Sub Form_Load()

    '***** fase di caricamento...
    loading = True

    '***** Mappatura ingressi analogici
    CaricaTabellaIngressiAnalogici
    
    '***** Mappatura ingressi digitali
    CaricaTabellaIngressiDigitali

    '***** Mappatura uscite analogiche
    CaricaTabellaUsciteAnalogiche

    '***** Mappatura uscite digitali
    CaricaTabellaUsciteDigitali
    
    '***** Linea *****
    CaricaLinea
    CaricaImpostazioni
    
    '***** termine fase di caricamento...
    loading = False
    cmd_salva.Enabled = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call SalvaModifiche
    Form1.Timer1.Enabled = True

End Sub

Private Sub SalvaModifiche()

    Dim salva As Boolean
    
    If MappaturaAImodificata Then salva = True
    If MappaturaDImodificata Then salva = True
    If MappaturaAOmodificata Then salva = True
    If MappaturaDOmodificata Then salva = True
    If LineaModificata Then salva = True
    If ImpostazioniModificate Then salva = True
    
    If salva Then
        If MsgBox("Salvo le ultime modifiche?", vbYesNo Or vbQuestion, "Mappatura I/O") = vbYes Then
            If MappaturaAImodificata Then SalvaMappaAI
            If MappaturaDImodificata Then SalvaMappaDI
            If MappaturaAOmodificata Then SalvaMappaAO
            If MappaturaDOmodificata Then SalvaMappaDO
            If LineaModificata Then SalvaLinea
            If ImpostazioniModificate Then SalvaImpostazioni
            
            Call MsgBox("Salvataggio effettuato", , "Mappatura I/O")
        End If
        cmd_salva.Enabled = False
    End If
    
    MappaturaAImodificata = False
    MappaturaDImodificata = False
    MappaturaAOmodificata = False
    MappaturaDOmodificata = False
    LineaModificata = False
    ImpostazioniModificate = False

End Sub

'***************************************************************************************
'                           SEZIONE MAPPATURA INGRESSI ANALOGICI                       *
'***************************************************************************************
Private Sub GridEX_AI_AfterUpdate()

    Dim i As Integer
    
    On Error GoTo GestErr
    'Parametri(0, riga) = Descrizione
    'Parametri(1, riga) = Numero DB
    'Parametri(2, riga) = Indirizzo base (OFFSET)
    'Parametri(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'Parametri(4, riga) = N. ingressi da leggere
    'Parametri(5, riga) = N. bytes da leggere (valore calcolato)
    'Parametri(6, riga) = Range Ingressi (valore calcolato)
    
    If GridEX_AI.Row = -1 Then
        GridEX_AI.Rebind
    Else
        For i = 0 To mRecordsetCols - 1
            Parametri(i, GridEX_AI.Row - 1) = GridEX_AI.Value(i + 1)
        Next
        
        'calcola N. bytes
        Select Case Parametri(3, GridEX_AI.Row - 1)
            Case "0"  'intero: 2 bytes per lettura
                Parametri(5, GridEX_AI.Row - 1) = Val(Parametri(4, GridEX_AI.Row - 1)) * 2
            Case Else 'float: 4 bytes per lettura
                Parametri(5, GridEX_AI.Row - 1) = Val(Parametri(4, GridEX_AI.Row - 1)) * 4
        End Select
        
        Call RefreshRangeAI
        GridEX_AI.Rebind
        
    End If
    
    MappaturaAImodificata = True
    cmd_salva.Enabled = True
    
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AI_AfterUpdate(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AI_AfterUpdate(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub GridEX_AI_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    
    Dim i As Long
    Dim j As Long

    On Error GoTo GestErr
    'First shift the rows
    For i = RowIndex - 1 To mRecordCount - 2
        For j = 0 To mRecordsetCols
            Parametri(j, i) = Parametri(j, i + 1)
        Next
    Next
    
    'decrement rowcount and redim array
    mRecordCount = mRecordCount - 1
    If mRecordCount > 0 Then ReDim Preserve Parametri(0 To mRecordsetCols, 0 To mRecordCount - 1)
    Call RefreshRangeAI
    
    MappaturaAImodificata = True
    cmd_salva.Enabled = True

    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AI_UnboundDelete(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AI_UnboundDelete(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub GridEX_AI_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)

    If MsgBox("Confermare cancellazione dati.", vbOKCancel Or vbExclamation, "Mappatura AI") <> vbOK Then Cancel = True
    
End Sub

Private Sub GridEX_ai_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer
    Dim n_bytes As Integer
    Dim bytes_dato As Integer

    On Error GoTo GestErr
    
    'Increase the record count variable
    mRecordCount = mRecordCount + 1
    
    'redim the array holding the grid values
    If mRecordCount = 1 Then
        ReDim Parametri(0 To mRecordsetCols, 0 To mRecordCount - 1)
    Else
        ReDim Preserve Parametri(0 To mRecordsetCols, 0 To mRecordCount - 1)
    End If
    
    'write the new values in the last record
    For i = 1 To Values.ColCount - 2
        Parametri(i - 1, mRecordCount - 1) = Values(i)
    Next
    
    'calcola N. bytes
    Select Case Parametri(3, mRecordCount - 1)
        Case "0"  'intero: 2 bytes per lettura
            Parametri(5, mRecordCount - 1) = Val(Parametri(4, mRecordCount - 1)) * 2
        Case Else 'float: 4 bytes per lettura
            Parametri(5, mRecordCount - 1) = Val(Parametri(4, mRecordCount - 1)) * 4
    End Select

    Call RefreshRangeAI
    

    Exit Sub
    
GestErr:
    Debug.Print "GridEX_ai_UnboundAddNew(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_ai_UnboundAddNew(): " & Err.Description)
    Resume Next


End Sub

Private Sub GridEX_AI_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer

    On Error GoTo GestErr
    
    'Parametri(0, riga) = Descrizione
    'Parametri(1, riga) = Numero DB
    'Parametri(2, riga) = Indirizzo base
    'Parametri(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'Parametri(4, riga) = N. ingressi da leggere
    'Parametri(5, riga) = N. bytes da leggere (valore calcolato)
    'Parametri(6, riga) = Range Ingressi (valore calcolato)
    
    'set the field values
    'Note: Values array is 1-based and Parametri array is 0-based
    For i = 1 To Values.ColCount
        If i = 6 Then
            Values(i) = Val(Parametri(i - 1, RowIndex - 1))
        Else
            Values(i) = Parametri(i - 1, RowIndex - 1)
        End If
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_ai_UnboundReadData(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_ai_UnboundReadData(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub RefreshRangeAI()

    Dim i As Integer
    Dim range As String
    Dim n_ai As Integer
    Dim last_n_ai As Integer
    Dim tot_ai As Integer
    
    On Error GoTo GestErr
    
    'Parametri(0, riga) = Descrizione
    'Parametri(1, riga) = Numero DB
    'Parametri(2, riga) = Indirizzo base
    'Parametri(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'Parametri(4, riga) = N. ingressi da leggere
    'Parametri(5, riga) = N. bytes da leggere (valore calcolato)
    'Parametri(6, riga) = Range Ingressi (valore calcolato)
    
    For i = 0 To mRecordCount - 1
        
        If range = "" Then
            n_ai = Val(Parametri(4, i))
            tot_ai = n_ai
            range = "0 - " & (n_ai - 1)
        Else
            n_ai = Val(Parametri(4, i))
            range = tot_ai & " - " & (tot_ai + n_ai - 1)
            tot_ai = tot_ai + n_ai
        End If
        Parametri(6, i) = range
        
    Next i
    
    Exit Sub
    
GestErr:
    Debug.Print "RefreshRangeAI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "RefreshRangeAI(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub SalvaMappaAI()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    
    On Error GoTo GestErr
    
    hdl = FreeFile
    Open App.Path & "\MappaAI.ini" For Output As #hdl
    
    GridEX_AI.MoveFirst
    For i = 0 To mRecordCount - 1
        riga_dati = ""
        
        For col = 0 To 4
            If col < 4 Then
                riga_dati = riga_dati & GridEX_AI.Value(col + 1) & ";"
            Else
                riga_dati = riga_dati & GridEX_AI.Value(col + 1)
            End If
        Next col
        
        Print #hdl, riga_dati
        GridEX_AI.MoveNext
        
    Next i
    Close (hdl)
    
    Exit Sub
    
GestErr:
    Debug.Print "SalvaMappaAI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaMappaAI(): " & Err.Description)
    Resume Next
    
End Sub

Public Sub CaricaMappaAI()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    Dim dati() As String
    
    On Error GoTo GestErr
    
    ReDim Parametri(0 To 6, 0 To 0)
    mRecordsetCols = UBound(Parametri, 1)
    mRecordCount = 0

    If Dir(App.Path & "\MappaAI.ini") = "" Then Exit Sub
    
    hdl = FreeFile
    Open App.Path & "\MappaAI.ini" For Input As #hdl
    Do While Not EOF(hdl)
        
        Line Input #hdl, riga_dati
        dati = Split(riga_dati, ";")
        
        mRecordCount = mRecordCount + 1
        ReDim Preserve Parametri(0 To mRecordsetCols, 0 To mRecordCount - 1)
        
        For col = 0 To UBound(dati, 1)
            Parametri(col, mRecordCount - 1) = dati(col)
        Next col
        
        'calcola N. bytes
        'Matteo febbraio 2016 - Corretto indice per matrice parametri
        Select Case Parametri(3, mRecordCount - 1)
            Case "0"  'intero: 2 bytes per lettura
                Parametri(5, mRecordCount - 1) = Val(Parametri(4, mRecordCount - 1)) * 2
            Case Else 'float: 4 bytes per lettura
                Parametri(5, mRecordCount - 1) = Val(Parametri(4, mRecordCount - 1)) * 4
        End Select
        
    Loop
    
    Close (hdl)
    
    Call RefreshRangeAI
    
    '********************************************
    For i = 0 To mRecordCount - 1
        For col = 0 To 6 'UBound(dati, 1)
            MappaturaAI(i, col) = Parametri(col, i)
        Next col
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "CaricaMappaAI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaMappaAI(): " & Err.Description)
    Resume Next

    
End Sub


'***************************************************************************************
'                           SEZIONE MAPPATURA INGRESSI DIGITALI                        *
'***************************************************************************************
Private Sub GridEX_DI_AfterUpdate()

    On Error GoTo GestErr
    
    If GridEX_DI.Row = -1 Then
        GridEX_DI.Rebind
    Else
        For i = 0 To mRecordsetColsDI - 1
            ParametriDI(i, GridEX_DI.Row - 1) = GridEX_DI.Value(i + 1)
        Next
        
        'calcola N. ingressi
        ParametriDI(5, GridEX_DI.Row - 1) = Val(ParametriDI(3, GridEX_DI.Row - 1)) * 8
        
        Call RefreshRangeDI
        GridEX_DI.Rebind
    
    End If
    
    MappaturaDImodificata = True
    cmd_salva.Enabled = True

    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DI_AfterUpdate(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DI_AfterUpdate(): " & Err.Description)
    Resume Next

End Sub

Private Sub GridEX_DI_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)

    If MsgBox("Confermare cancellazione dati.", vbOKCancel Or vbExclamation, "Mappatura DI") <> vbOK Then Cancel = True
    
End Sub

Private Sub GridEX_DI_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer
    Dim n_bytes As Integer
    Dim bytes_dato As Integer

    On Error GoTo GestErr
    
    'Increase the record count variable
    mRecordCountDI = mRecordCountDI + 1
    
    'redim the array holding the grid values
    If mRecordCountDI = 1 Then
        ReDim ParametriDI(0 To mRecordsetColsDI, 0 To mRecordCountDI - 1)
    Else
        ReDim Preserve ParametriDI(0 To mRecordsetColsDI, 0 To mRecordCountDI - 1)
    End If
    
    'write the new values in the last record
    For i = 1 To Values.ColCount - 2
        ParametriDI(i - 1, mRecordCountDI - 1) = Values(i)
    Next
    
    'calcola N. ingressi
    ParametriDI(5, mRecordCountDI - 1) = Val(ParametriDI(3, mRecordCountDI - 1)) * 8
    
    Call RefreshRangeDI
    
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DI_UnboundAddNew(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DI_UnboundAddNew(): " & Err.Description)
    Resume Next


End Sub

Private Sub GridEX_DI_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    
    Dim i As Long
    Dim j As Long

    'First shift the rows
    For i = RowIndex - 1 To mRecordCountDI - 2
        For j = 0 To mRecordsetColsDI
            ParametriDI(j, i) = ParametriDI(j, i + 1)
        Next
    Next
    
    'decrement rowcount and redim array
    mRecordCountDI = mRecordCountDI - 1
    If mRecordCountDI > 0 Then ReDim Preserve ParametriDI(0 To mRecordsetColsDI, 0 To mRecordCountDI - 1)
    Call RefreshRangeDI
    
    MappaturaDImodificata = True
    cmd_salva.Enabled = True

    On Error GoTo GestErr
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DI_UnboundDelete(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DI_UnboundDelete(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub GridEX_DI_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer

    On Error GoTo GestErr
    
    'set the field values
    'Note: Values array is 1-based and Parametri array is 0-based

    For i = 1 To Values.ColCount
        If i = 5 Then
            Values(i) = Val(ParametriDI(i - 1, RowIndex - 1))
        Else
            Values(i) = ParametriDI(i - 1, RowIndex - 1)
        End If
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DI_UnboundReadData(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DI_UnboundReadData(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub RefreshRangeDI()

    Dim i As Integer
    Dim range As String
    Dim n_di As Integer
    Dim last_n_di As Integer
    Dim tot_di As Integer
    
    On Error GoTo GestErr
    
    For i = 0 To mRecordCountDI - 1
        
        If range = "" Then
            n_di = Val(ParametriDI(5, i))
            tot_di = n_di
            range = "0 - " & (tot_di - 1)
        Else
            n_di = Val(ParametriDI(5, i))
            range = tot_di & " - " & (tot_di + n_di - 1)
            tot_di = tot_di + n_di
        End If
        ParametriDI(6, i) = range
        
    Next i
    
    Exit Sub
    
GestErr:
    Debug.Print "RefreshRangeDI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "RefreshRangeDI(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub SalvaMappaDI()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    
    On Error GoTo GestErr
    
    hdl = FreeFile
    Open App.Path & "\MappaDI.ini" For Output As #hdl
    
    GridEX_DI.MoveFirst
    For i = 0 To mRecordCountDI - 1
        riga_dati = ""
        
        For col = 0 To 4
            If col < 4 Then
                riga_dati = riga_dati & GridEX_DI.Value(col + 1) & ";"
            Else
                riga_dati = riga_dati & GridEX_DI.Value(col + 1)
            End If
        Next col
        
        Print #hdl, riga_dati
        GridEX_DI.MoveNext
        
    Next i
    Close (hdl)
    
    Exit Sub
    
GestErr:
    Debug.Print "SalvaMappaDI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaMappaDI(): " & Err.Description)
    Resume Next

    
End Sub

Public Sub CaricaMappaDI()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    Dim dati() As String
    
    On Error GoTo GestErr
    
    ReDim ParametriDI(0 To 6, 0 To 0)
    mRecordsetColsDI = UBound(ParametriDI, 1)
    mRecordCountDI = 0

    If Dir(App.Path & "\MappaDI.ini") = "" Then Exit Sub
    
    hdl = FreeFile
    Open App.Path & "\MappaDI.ini" For Input As #hdl
    Do While Not EOF(hdl)
        
        Line Input #hdl, riga_dati
        dati = Split(riga_dati, ";")
        
        mRecordCountDI = mRecordCountDI + 1
        ReDim Preserve ParametriDI(0 To mRecordsetColsDI, 0 To mRecordCountDI - 1)
        
        For col = 0 To UBound(dati, 1)
            ParametriDI(col, mRecordCountDI - 1) = dati(col)
        Next col
        
        'calcola N. ingressi
        ParametriDI(5, mRecordCountDI - 1) = Val(ParametriDI(3, mRecordCountDI - 1)) * 8

    Loop
    
    Close (hdl)
    
    Call RefreshRangeDI
    
    '********************************************
    For i = 0 To mRecordCountDI - 1
        For col = 0 To UBound(dati, 1)
            MappaturaDI(i, col) = ParametriDI(col, i)
        Next col
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "CaricaMappaDI(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaMappaDI(): " & Err.Description)
    Resume Next

    
End Sub


'***************************************************************************************
'                           SEZIONE MAPPATURA USCITE ANALOGICHE                        *
'***************************************************************************************
Private Sub GridEX_AO_AfterUpdate()

    Dim i As Integer
    
    On Error GoTo GestErr
    'ParametriAO(0, riga) = Descrizione
    'ParametriAO(1, riga) = Numero DB
    'ParametriAO(2, riga) = Indirizzo base (OFFSET)
    'ParametriAO(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'ParametriAO(4, riga) = N. uscite da scrivere
    'ParametriAO(5, riga) = N. bytes da scrivere (valore calcolato)
    'ParametriAO(6, riga) = Range uscite (valore calcolato)
    
    If GridEX_AO.Row = -1 Then
        GridEX_AO.Rebind
    Else
        For i = 0 To mRecordsetColsAO - 1
            ParametriAO(i, GridEX_AO.Row - 1) = GridEX_AO.Value(i + 1)
        Next
        
        'calcola N. ingressi
        Select Case ParametriAO(3, GridEX_AO.Row - 1)
            Case "0"
                ParametriAO(5, GridEX_AO.Row - 1) = Val(ParametriAO(4, GridEX_AO.Row - 1)) * 2
            Case Else
                ParametriAO(5, GridEX_AO.Row - 1) = Val(ParametriAO(4, GridEX_AO.Row - 1)) * 4
        End Select
    
        Call RefreshRangeAO
        GridEX_AO.Rebind
        
    End If
    
    MappaturaAOmodificata = True
    cmd_salva.Enabled = True
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AO_AfterUpdate(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AO_AfterUpdate(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub GridEX_AO_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    Dim i As Long
    Dim j As Long

    On Error GoTo GestErr
    'First shift the rows
    For i = RowIndex - 1 To mRecordCountAO - 2
        For j = 0 To mRecordsetColsAO
            ParametriAO(j, i) = ParametriAO(j, i + 1)
        Next
    Next
    
    'decrement rowcount and redim array
    mRecordCountAO = mRecordCountAO - 1
    If mRecordCountAO > 0 Then ReDim Preserve ParametriAO(0 To mRecordsetColsAO, 0 To mRecordCountAO - 1)
    Call RefreshRangeAO
    
    MappaturaAOmodificata = True
    cmd_salva.Enabled = True
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AO_UnboundDelete(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AO_UnboundDelete(): " & Err.Description)
    Resume Next

End Sub

Private Sub GridEX_AO_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)

    If MsgBox("Confermare cancellazione dati.", vbOKCancel Or vbExclamation, "Mappatura AI") <> vbOK Then Cancel = True
    
End Sub

Private Sub GridEX_AO_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer
    Dim n_bytes As Integer
    Dim bytes_dato As Integer

    On Error GoTo GestErr
    
    'ParametriAO(0, riga) = Descrizione
    'ParametriAO(1, riga) = Numero DB
    'ParametriAO(2, riga) = Indirizzo base (OFFSET)
    'ParametriAO(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'ParametriAO(4, riga) = N. uscite da scrivere
    'ParametriAO(5, riga) = N. bytes da scrivere (valore calcolato)
    'ParametriAO(6, riga) = Range uscite (valore calcolato)
    
    'Increase the record count variable
    mRecordCountAO = mRecordCountAO + 1
    'redim the array holding the grid values
    If mRecordCountAO = 1 Then
        ReDim ParametriAO(0 To mRecordsetColsAO, 0 To mRecordCountAO - 1)
    Else
        ReDim Preserve ParametriAO(0 To mRecordsetColsAO, 0 To mRecordCountAO - 1)
    End If
    
    'write the new values in the last record
    For i = 1 To Values.ColCount - 2
        ParametriAO(i - 1, mRecordCountAO - 1) = Values(i)
    Next
    
    'calcola N. ingressi
    Select Case ParametriAO(3, mRecordCountAO - 1)
        Case "0"
            ParametriAO(5, mRecordCountAO - 1) = Val(ParametriAO(4, mRecordCountAO - 1)) * 2
        Case Else
            ParametriAO(5, mRecordCountAO - 1) = Val(ParametriAO(4, mRecordCountAO - 1)) * 4
    End Select

    Call RefreshRangeAO
    

    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AO_UnboundAddNew(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AO_UnboundAddNew(): " & Err.Description)
    Resume Next


End Sub

Private Sub GridEX_AO_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer

    On Error GoTo GestErr
    
    'ParametriAO(0, riga) = Descrizione
    'ParametriAO(1, riga) = Numero DB
    'ParametriAO(2, riga) = Indirizzo base (OFFSET)
    'ParametriAO(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'ParametriAO(4, riga) = N. uscite da scrivere
    'ParametriAO(5, riga) = N. bytes da scrivere (valore calcolato)
    'ParametriAO(6, riga) = Range uscite (valore calcolato)
   
    'set the field values
    'Note: Values array is 1-based and Parametri array is 0-based

    For i = 1 To Values.ColCount
        If i = 6 Then
            Values(i) = Val(ParametriAO(i - 1, RowIndex - 1))
        Else
            Values(i) = ParametriAO(i - 1, RowIndex - 1)
        End If
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_AO_UnboundReadData(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_AO_UnboundReadData(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub RefreshRangeAO()

    Dim i As Integer
    Dim range As String
    Dim n_ao As Integer
    Dim last_n_ao As Integer
    Dim tot_ao As Integer
    
    'ParametriAO(0, riga) = Descrizione
    'ParametriAO(1, riga) = Numero DB
    'ParametriAO(2, riga) = Indirizzo base (OFFSET)
    'ParametriAO(3, riga) = Tipo variabile: 0=intero (2 bytes)   1=float (4 bytes)
    'ParametriAO(4, riga) = N. uscite da scrivere
    'ParametriAO(5, riga) = N. bytes da scrivere (valore calcolato)
    'ParametriAO(6, riga) = Range uscite (valore calcolato)
    
    On Error GoTo GestErr
    
    For i = 0 To mRecordCountAO - 1
        
        If range = "" Then
            n_ao = Val(ParametriAO(4, i))
            tot_ao = n_ao
            range = "0 - " & (n_ao - 1)
        Else
            n_ao = Val(ParametriAO(4, i))
            range = tot_ao & " - " & (tot_ao + n_ao - 1)
            tot_ao = tot_ao + n_ao
        End If
        ParametriAO(6, i) = range
        
    Next i
    
    Exit Sub
    
GestErr:
    Debug.Print "RefreshRangeAO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "RefreshRangeAO(): " & Err.Description)
    Resume Next
    
End Sub

Private Sub SalvaMappaAO()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    
    On Error GoTo GestErr
    
    hdl = FreeFile
    Open App.Path & "\MappaAO.ini" For Output As #hdl
    
    GridEX_AO.MoveFirst
    For i = 0 To mRecordCountAO - 1
        riga_dati = ""
        
        For col = 0 To 4
            If col < 4 Then
                riga_dati = riga_dati & GridEX_AO.Value(col + 1) & ";"
            Else
                riga_dati = riga_dati & GridEX_AO.Value(col + 1)
            End If
        Next col
        
        Print #hdl, riga_dati
        GridEX_AO.MoveNext
        
    Next i
    Close (hdl)
    
    Exit Sub
    
GestErr:
    Debug.Print "SalvaMappaAO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaMappaAO(): " & Err.Description)
    Resume Next

    
End Sub

Public Sub CaricaMappaAO()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    Dim dati() As String
    
    On Error GoTo GestErr
    
    ReDim ParametriAO(0 To 6, 0 To 0)
    mRecordsetColsAO = UBound(Parametri, 1)
    mRecordCountAO = 0

    If Dir(App.Path & "\MappaAO.ini") = "" Then Exit Sub
    
    hdl = FreeFile
    Open App.Path & "\MappaAO.ini" For Input As #hdl
    Do While Not EOF(hdl)
        
        Line Input #hdl, riga_dati
        dati = Split(riga_dati, ";")
        
        mRecordCountAO = mRecordCountAO + 1
        ReDim Preserve ParametriAO(0 To mRecordsetColsAO, 0 To mRecordCountAO - 1)
        
        For col = 0 To UBound(dati, 1)
            ParametriAO(col, mRecordCountAO - 1) = dati(col)
        Next col
        
        'calcola N. ingressi
        Select Case ParametriAO(3, mRecordCountAO - 1)
            Case "0"
                ParametriAO(5, mRecordCountAO - 1) = Val(ParametriAO(4, mRecordCountAO - 1)) * 2
            Case Else
                ParametriAO(5, mRecordCountAO - 1) = Val(ParametriAO(4, mRecordCountAO - 1)) * 4
        End Select

        
    Loop
    
    Close (hdl)
    
    Call RefreshRangeAO
    
    '********************************************
    For i = 0 To mRecordCountAO - 1
        For col = 0 To 6  'UBound(dati, 1)
            MappaturaAO(i, col) = ParametriAO(col, i)
        Next col
    Next
    
    Exit Sub
    
GestErr:
    Debug.Print "CaricaMappaAO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaMappaAO(): " & Err.Description)
    Resume Next

    
End Sub


'***************************************************************************************
'                           SEZIONE MAPPATURA USCITE DIGITALI                          *
'***************************************************************************************
Private Sub GridEX_DO_AfterUpdate()

    On Error GoTo GestErr

    If GridEX_DO.Row = -1 Then
        GridEX_DO.Rebind
    Else
        For i = 0 To mRecordsetColsDO - 1
            ParametriDO(i, GridEX_DO.Row - 1) = GridEX_DO.Value(i + 1)
        Next

        'calcola N. uscite
        ParametriDO(5, GridEX_DO.Row - 1) = Val(ParametriDO(3, GridEX_DO.Row - 1)) * 8

        Call RefreshRangeDO
        GridEX_DO.Rebind

    End If

    MappaturaDOmodificata = True
    cmd_salva.Enabled = True
    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DO_AfterUpdate(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DO_AfterUpdate(): " & Err.Description)
    Resume Next

End Sub

Private Sub GridEX_DO_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)

    If MsgBox("Confermare cancellazione dati.", vbOKCancel Or vbExclamation, "Mappatura DI") <> vbOK Then Cancel = True

End Sub

Private Sub GridEX_DO_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer
    Dim n_bytes As Integer
    Dim bytes_dato As Integer

    On Error GoTo GestErr

    'Increase the record count variable
    mRecordCountDO = mRecordCountDO + 1

    'redim the array holding the grid values
    If mRecordCountDO = 1 Then
        ReDim ParametriDO(0 To mRecordsetColsDO, 0 To mRecordCountDO - 1)
    Else
        ReDim Preserve ParametriDO(0 To mRecordsetColsDO, 0 To mRecordCountDO - 1)
    End If

    'write the new values in the last record
    For i = 1 To Values.ColCount - 2
        ParametriDO(i - 1, mRecordCountDO - 1) = Values(i)
    Next

    'calcola N. uscite
    ParametriDO(5, mRecordCountDO - 1) = Val(ParametriDO(3, mRecordCountDO - 1)) * 8

    Call RefreshRangeDO

    Exit Sub

GestErr:
    Debug.Print "GridEX_DO_UnboundAddNew(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DO_UnboundAddNew(): " & Err.Description)
    Resume Next


End Sub

Private Sub GridEX_DO_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)

    Dim i As Long
    Dim j As Long

    On Error GoTo GestErr
    'First shift the rows
    For i = RowIndex - 1 To mRecordCountDO - 2
        For j = 0 To mRecordsetColsDO
            ParametriDO(j, i) = ParametriDO(j, i + 1)
        Next
    Next

    'decrement rowcount and redim array
    mRecordCountDO = mRecordCountDO - 1
    If mRecordCountDO > 0 Then ReDim Preserve ParametriDO(0 To mRecordsetColsDO, 0 To mRecordCountDO - 1)
    Call RefreshRangeDO

    MappaturaDOmodificata = True
    cmd_salva.Enabled = True

    Exit Sub
    
GestErr:
    Debug.Print "GridEX_DO_UnboundDelete(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DO_UnboundDelete(): " & Err.Description)
    Resume Next

End Sub

Private Sub GridEX_DO_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Dim i As Integer

    On Error GoTo GestErr

    'set the field values
    'Note: Values array is 1-based and Parametri array is 0-based

    For i = 1 To Values.ColCount
        If i = 5 Then
            Values(i) = Val(ParametriDO(i - 1, RowIndex - 1))
        Else
            Values(i) = ParametriDO(i - 1, RowIndex - 1)
        End If
    Next

    Exit Sub

GestErr:
    Debug.Print "GridEX_DO_UnboundReadData(): " & Err.Description
    Call SalvaLog(ERR_LOG, "GridEX_DO_UnboundReadData(): " & Err.Description)
    Resume Next

End Sub

Private Sub RefreshRangeDO()

    Dim i As Integer
    Dim range As String
    Dim n_do As Integer
    Dim last_n_do As Integer
    Dim tot_do As Integer

    On Error GoTo GestErr

    For i = 0 To mRecordCountDO - 1

        If range = "" Then
            n_do = Val(ParametriDO(5, i))
            tot_do = n_do
            range = "0 - " & (tot_do - 1)
        Else
            n_do = Val(ParametriDO(5, i))
            range = tot_do & " - " & (tot_do + n_do - 1)
            tot_do = tot_do + n_do
        End If
        ParametriDO(6, i) = range

    Next i

    Exit Sub

GestErr:
    Debug.Print "RefreshRangeDO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "RefreshRangeDO(): " & Err.Description)
    Resume Next

End Sub

Public Sub CaricaImpostazioni()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    Dim par() As String


    On Error GoTo GestErr

    If Dir(App.Path & "\Impostazioni.ini") = "" Then Exit Sub

    hdl = FreeFile
    Open App.Path & "\Impostazioni.ini" For Input As #hdl
    
    Do While Not EOF(hdl)
        Line Input #hdl, riga_dati
        If InStr(riga_dati, "=") > 0 Then
            par = Split(riga_dati, "=")
            Select Case UCase(Trim(par(0)))
            
                Case "ALARM_ENABLED"
                    chk_allarmi.Value = IIf(CBool(Trim(par(1))), 1, 0)
                    bComunicationAlarmEnabled = IIf(chk_allarmi.Value = 1, True, False)

                Case "ALARM_INDEX"
                    txt_morsetto_allarme.Text = Trim(par(1))
                    iComunicationAlarmIndex = Val(Trim(par(1)))
                    
                Case "BF_COMUNICATOR"
                    chk_bfcomunicator.Value = IIf(CBool(Trim(par(1))), 1, 0)
                    BFComunicatorDisabled = IIf(chk_bfcomunicator.Value = 0, True, False)
                    
                Case "WD_ENABLED"
                    chk_watchdog.Value = IIf(CBool(Trim(par(1))), 1, 0)
                    bWatchDogEnabled = chk_watchdog.Value
                    pic_watchdog.Visible = chk_watchdog.Value

                Case "WD_TAG"
                    txt_tag_wd.Text = Trim(par(1))
                    sWD_TAG = Trim(par(1))
                    
            End Select
        End If
        
    Loop
    Close (hdl)

    ImpostazioniModificate = False

    Exit Sub

GestErr:
    Debug.Print "CaricaImpostazioni(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaImpostazioni(): " & Err.Description)
    Resume Next

End Sub

Public Sub SalvaImpostazioni()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String

    On Error GoTo GestErr

    hdl = FreeFile
    Open App.Path & "\Impostazioni.ini" For Output As #hdl

    Print #hdl, "ALARM_ENABLED = " & chk_allarmi.Value
    Print #hdl, "ALARM_INDEX = " & Val(txt_morsetto_allarme.Text)
    Print #hdl, "BF_COMUNICATOR = " & chk_bfcomunicator.Value
    Print #hdl, "WD_ENABLED = " & chk_watchdog.Value
    Print #hdl, "WD_TAG = " & txt_tag_wd.Text
    
    Close (hdl)
    
    BFComunicatorDisabled = IIf(chk_bfcomunicator.Value = 0, True, False)
    
    Exit Sub

GestErr:
    Debug.Print "SalvaImpostazioni: " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaImpostazioni(): " & Err.Description)
    Resume Next

End Sub

Private Sub SettaModifiche()
    If Not loading Then
        ImpostazioniModificate = True
        cmd_salva.Enabled = True
    End If
End Sub

Private Sub SalvaMappaDO()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String

    On Error GoTo GestErr

    hdl = FreeFile
    Open App.Path & "\MappaDO.ini" For Output As #hdl

    GridEX_DO.MoveFirst
    For i = 0 To mRecordCountDO - 1
        riga_dati = ""

        For col = 0 To 4
            If col < 4 Then
                riga_dati = riga_dati & GridEX_DO.Value(col + 1) & ";"
            Else
                riga_dati = riga_dati & GridEX_DO.Value(col + 1)
            End If
        Next col

        Print #hdl, riga_dati
        GridEX_DO.MoveNext

    Next i
    Close (hdl)

    Exit Sub

GestErr:
    Debug.Print "SalvaMappaDO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaMappaDO(): " & Err.Description)
    Resume Next


End Sub

Private Sub SalvaLinea()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String

    On Error GoTo GestErr

    hdl = FreeFile
    Open App.Path & "\Linea.ini" For Output As #hdl

    riga_dati = cmbLinea.ListIndex

    Print #hdl, riga_dati

    Close (hdl)

    Exit Sub

GestErr:
    Debug.Print "SalvaLinea(): " & Err.Description
    Call SalvaLog(ERR_LOG, "SalvaLinea(): " & Err.Description)
    Resume Next

End Sub
Public Sub CaricaMappaDO()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String
    Dim dati() As String

    On Error GoTo GestErr

    ReDim ParametriDO(0 To 6, 0 To 0)
    mRecordsetColsDO = UBound(ParametriDO, 1)
    mRecordCountDO = 0

    If Dir(App.Path & "\MappaDO.ini") = "" Then Exit Sub

    hdl = FreeFile
    Open App.Path & "\MappaDO.ini" For Input As #hdl
    Do While Not EOF(hdl)

        Line Input #hdl, riga_dati
        dati = Split(riga_dati, ";")

        mRecordCountDO = mRecordCountDO + 1
        ReDim Preserve ParametriDO(0 To mRecordsetColsDO, 0 To mRecordCountDO - 1)

        For col = 0 To UBound(dati, 1)
            ParametriDO(col, mRecordCountDO - 1) = dati(col)
        Next col

        'calcola N. uscite
        ParametriDO(5, mRecordCountDO - 1) = Val(ParametriDO(3, mRecordCountDO - 1)) * 8

    Loop

    Close (hdl)

    Call RefreshRangeDO

    '********************************************
    For i = 0 To mRecordCountDO - 1
        For col = 0 To UBound(dati, 1)
            MappaturaDO(i, col) = ParametriDO(col, i)
        Next col
    Next

    Exit Sub

GestErr:
    Debug.Print "CaricaMappaDO(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaMappaDO(): " & Err.Description)
    Resume Next


End Sub
Public Sub CaricaLinea()

    Dim i As Integer
    Dim hdl As Integer
    Dim col As Integer
    Dim riga_dati As String


    On Error GoTo GestErr

    LineaBFlab = 1
    If Dir(App.Path & "\Linea.ini") = "" Then Exit Sub

    hdl = FreeFile
    Open App.Path & "\Linea.ini" For Input As #hdl
    
    If Not EOF(hdl) Then
        Line Input #hdl, riga_dati
        LineaBFlab = Val(riga_dati)
    End If

    Close (hdl)

    cmbLinea.ListIndex = LineaBFlab
    
    LineaModificata = False
    
    Exit Sub

GestErr:
    Debug.Print "CaricaLinea(): " & Err.Description
    Call SalvaLog(ERR_LOG, "CaricaLinea(): " & Err.Description)
    Resume Next

End Sub

Private Sub txt_morsetto_allarme_Change()

  If Val(txt_morsetto_allarme.Text) > 255 Then txt_morsetto_allarme.Text = "255"
  
  Call SettaModifiche
  
End Sub

Private Sub txt_tag_wd_Change()

  Call SettaModifiche
  
End Sub
