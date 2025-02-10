Attribute VB_Name = "GestioneDatabase"
Option Explicit

Dim rsDB As Object

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 1. Deve essere creato un allarme nelle digitali di linea "CONNESSIONE DATABASE" con morsetto 99.01 (?)
' 2. Deve essere creata una procedura che testa la connessione al database
' 3. E' possibile sfruttare la procedura anche all'avvio del software senza mettere i ritardi?
' 4. L'anomalia deve rientrare quando la connessione viene ristabilita.
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub CheckDBConnection()

    Const query = "SELECT * FROM wds_gentab"

    On Error GoTo GestErrore
    
    ConnessioneValida = False
    NewDataObj rsDB
    rsDB.selectionfast query
    
    manValoreDigitale(999, 2) = 0
    
    Set rsDB = Nothing
    
    ConnessioneValida = True
    Exit Sub
    
GestErrore:
    Select Case Err()
        Case -2147467259    'SQL inesistente
            Call WindasLog("CheckDBConnection: Connessione non disponibile!", 1, "OPC")
            
        Case Else
            Call WindasLog("CheckDBConnection: " & Error(Err()), 1, "OPC")
    End Select
   manValoreDigitale(999, 2) = 1

End Sub

Sub GetConnectionParam()
'***** Lettura parametri di configurazione database su file bfdesk.xml *****

    Dim AccessObj As Object
    Dim Crypt As Object
    Dim iTempDB As Integer
    Dim TempDB(99) As ConnectionsDB
    Dim iConn As Integer
    Dim FileXML As String
        
    On Error GoTo GestErrore
    
    FileXML = App.Path & "\connections.xml"
    
    Set AccessObj = CreateObject("AttimoFwk.CData")
    With AccessObj
        .XML_OpenSchema (CStr(FileXML))
        iTempDB = 0
        Do While Not .iseof

            '** Verifica della presenza del CodiceStazione
            TempDB(iTempDB).StationCode = .XML_ReadField("Stationcode")
            If Len(TempDB(iTempDB).StationCode) = 0 Then
                MsgBox "Codice Stazione non configurato! Esecuzione terminata.", vbCritical
                End
            End If
            
            '** Estrazione dei dati di ogni DB
            TempDB(iTempDB).AppServer = .XML_ReadField("SERVER")
            TempDB(iTempDB).AppDatabase = .XML_ReadField("DATABASE")
            TempDB(iTempDB).AppDBType = .XML_ReadField("DBTYPE")
            TempDB(iTempDB).AppDbVersion = .XML_ReadField("CNMODE")
            TempDB(iTempDB).AppDBUser = .XML_ReadField("USER")
            TempDB(iTempDB).AppDBPwd = .XML_ReadField("PASSWORD")
            TempDB(iTempDB).AppDefaultDB = .XML_ReadField("DEFAULT")
            TempDB(iTempDB).AppOrderSAD = .XML_ReadField("OrdineCaricamentoDatiElementari")
            Set Crypt = CreateObject("AttimoFwk.CCrypt")
            TempDB(iTempDB).AppDBUser = Crypt.Decrypt(CStr(TempDB(iTempDB).AppDBUser))
            TempDB(iTempDB).AppDBPwd = Crypt.Decrypt(CStr(TempDB(iTempDB).AppDBPwd))
            Set Crypt = Nothing
            
            .MoveNext
            iTempDB = iTempDB + 1
        Loop
    End With
    Set AccessObj = Nothing
    
    '** Creazione dell'array delle connessioni
    iTempDB = iTempDB - 1
    ReDim Preserve connDB(iTempDB)
    For iConn = 0 To iTempDB
        connDB(iConn) = TempDB(iConn)
    Next iConn
    
    '** Ordino le connessioni
    Call DBConnectionsOrder
    
    '** Indice connessione di Default
    iConnDBDefault = DBConnectionsGetDefault()
    
    '** Codice Stazione
    gsClienteDi = connDB(iConnDBDefault).StationCode
    
    Exit Sub
    
GestErrore:
    Call WindasLog("GetConnectionParam " + Error(Err), 1, "OPC")

End Sub

Sub DBConnectionsOrder()

    Dim lngX As Long
    Dim lngY As Long
    Dim temp As ConnectionsDB
    
    On Error GoTo GestErrore
    
    For lngX = LBound(connDB) To (UBound(connDB) - 1)
        For lngY = LBound(connDB) To (UBound(connDB) - 1)
            If connDB(lngY).AppOrderSAD > connDB(lngY + 1).AppOrderSAD Then
                ' scambio gli elementi
                temp = connDB(lngY)
                connDB(lngY) = connDB(lngY + 1)
                connDB(lngY + 1) = temp
                
            End If
        Next lngY
    Next lngX
    
    Exit Sub
GestErrore:

    Call WindasLog("OrderDBConnections: " & Error(Err()), 1, "OPC")
    
End Sub

Function DBConnectionsGetDefault()
    
    Dim i As Integer
    
    On Error GoTo GestErrore
    
    For i = 0 To UBound(connDB)
        If connDB(i).AppDefaultDB Then
            DBConnectionsGetDefault = i 'salvo indice della connessione di defualt
            Exit For
        End If
    Next i
    
    Exit Function
GestErrore:

    Call WindasLog("DBConnectionsGetDefault: " & Error(Err()), 1, "OPC")
End Function

Public Sub NewDataObj(rs As Object, Optional idxConn As Integer = -1)

    On Error GoTo GestErrore
  
    'Nicola 29/11/2016
    'In base all'indice della connession alla quale voglio connettermi setto tutti i parametri sotto elencati
  
    If IsMissing(idxConn) Or idxConn < 0 Then 'se non passo indice di connessione
      idxConn = iConnDBDefault 'prendo quella di default caricata nel getconnectionparam
    End If
  
    Set rs = CreateObject("AttimoFwk.CData")
    With rs
      Select Case connDB(idxConn).AppDBType
        Case "ODBC"
          .SetDBType .Conn_ODBC
        Case "UDL"
          .UDLFile = App.Path & "\bflab7.udl"
          .SetDBType .Conn_UDL
        Case "ACCESS"
          .SetDBType .Conn_Jet
        Case "SQL"
          .SetDBType .Conn_SQL
        Case "ORACLE"
          .SetDBType .Conn_Oracle
        Case "MYSQL"
          .SetDBType .Conn_MYSQL
      Case "MYSQL5"
          .SetDBType (.Conn_MYSQL5)
      End Select
      .SetServer connDB(idxConn).AppServer
      .SetDatabase connDB(idxConn).AppDatabase, connDB(idxConn).AppDBUser, connDB(idxConn).AppDBPwd
      .SetDbversion connDB(idxConn).AppDbVersion
      .SetMessages False
      .Err_Activate = True
    End With
    
    Exit Sub
GestErrore:
    Call WindasLog("NewDataObj: " & Error(Err()), 1, "OPC")

End Sub

