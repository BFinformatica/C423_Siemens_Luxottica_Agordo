VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BFserver"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_form As Object
Dim m_Param As Variant
Dim Parametri() As String

Public Property Let Parameter(ByVal linea As Integer, ByVal code As String, ByVal Value As Variant)
    
    On Error GoTo GestErr
    
    Select Case UCase(code)
    
      Case "CONNECTION"
        
    End Select
    
    Exit Property
  
GestErr:
    warn_msg = "Let Parameter: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next
    
End Property

Public Property Get Parameter(ByVal linea As Integer, ByVal code As String) As Variant
    
    On Error GoTo GestErr
    
    Select Case UCase(code)
    
      Case "CONNECTION"
        
    End Select
    
    Exit Property
  
GestErr:
    warn_msg = "Get Parameter: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next
    
End Property

Public Property Get AnInput(ByVal index As Integer) As Integer
    
    On Error GoTo GestErr
    
    AnInput = AnalogReadings(index)
    
    Exit Property
  
GestErr:
    warn_msg = "Get AnInput: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Property

Public Property Get AnInputFromDCS(ByVal index As Integer) As Integer
    
    On Error GoTo GestErr
    
    AnInputFromDCS = AnalogFromDCS(index)
    
    Exit Property
  
GestErr:
    warn_msg = "Get AnInputFromDCS: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Property

Public Property Get DigInput(ByVal index As Integer) As Integer
    
    On Error GoTo GestErr
    
    DigInput = DigitalReadings(index)
    
    Exit Property
  
GestErr:
    warn_msg = "Get DigInput: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Property

Public Property Get DigInputFromDCS(ByVal index As Integer) As Integer
    
    On Error GoTo GestErr
    
    DigInputFromDCS = DigitalFromDCS(index)
    
    Exit Property
  
GestErr:
    warn_msg = "Get DigInput: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next

End Property

Public Property Get PlcConnected() As Boolean

    PlcConnected = PLC_Connected
    
End Property

Public Property Let AnOut(ByVal index As Integer, ByVal Value As Integer)

    iAnalogWritings(index) = Value

End Property

Public Property Let AnOutToDCS(ByVal index As Integer, ByVal Value As Single)

    AnalogToDCS(index) = Value

End Property

Public Property Let DigOut(ByVal index As Integer, ByVal Value As Integer)

    DigitalWritings(index) = Value

End Property

Public Property Let DigOutToDCS(ByVal index As Integer, ByVal Value As Integer)

    DigitalToDCS(index) = Value

End Property

Public Property Let SegnaliInUscita(ByVal Value As Boolean)
    
    NoDigOut = Value
  
End Property

Public Function MeasureData(code As Variant) As Variant

    On Error GoTo GestErr
    
    Select Case code
        Case "AI"
            MeasureData = AnalogReadings
        Case "DI"
            MeasureData = DigitalReadings
        Case "ONLINE"
            MeasureData = (no_com_ctr < TIMEOUT)
        Case "VERSION"
            MeasureData = str_version
        Case "WARNINGS"
            MeasureData = warn_msg
        Case Else
            MeasureData = Nothing
    
    End Select
    Exit Function
  
GestErr:
    warn_msg = "MeasureData: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next
    
End Function

Public Sub StartMe()

    Dim p As Integer
    Dim ndx As Integer
    
    On Error GoTo GestErr
    
    '*** michele - ping
    Set PingObj = CreateObject("AttimoFwk.CPing")

    str_version = "BFInformatica Driver per PLC Siemens S300-S400 versione " + Str(App.Major) + Str(App.Minor) + Str(App.Revision)
    
'    Form1.TmrInit.Enabled = False
    
    '**** Inizializza BFComunicator per trasmissione dati
    Call InizializzaComunicator
    
    If (Dir(App.Path & "\MappaAI.ini") = "") And (Dir(App.Path & "\MappaDI.ini") = "") Then
        Form2.Show
    Else
        Form2.CaricaLinea
        Form1.Show
    End If
    
    Exit Sub
  
GestErr:
    warn_msg = "StartMe: " & Err.Description
    Call SalvaLog(ERR_LOG, warn_msg)
    Resume Next
    
End Sub
Sub InizializzaComunicator()

    Dim iIdx As Integer
    
    'Alby ottobre 2012
    On Error GoTo GestErrore
    
    Set BFComunicator = GetObject("", "BFcomunicator.CloggerData")
    BFComunicator.StartDataSharing

    Exit Sub
    
GestErrore:
    Call SalvaLog(ERR_LOG, "InizializzaComunicator " + Error(Err))
    Resume fine
    
fine:


End Sub
Public Function CloseMe()
  
    Unload m_form
    Set m_form = Nothing
    
End Function

Public Function StartMeEseguito() As String

  StartMeEseguito = DriverLoaded
  
End Function



