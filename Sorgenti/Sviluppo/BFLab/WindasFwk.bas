Attribute VB_Name = "WindasFwk"
'luca aprile 2017 revisiono in seguito a modifica tabella di salvataggio log cambio configurazione
'Function AggiornaConfigurazione(NomeTabella As String, NomeColonna As String, Contenuto As String, clausolaWhere As String, Linea As Integer, Master_Slave As Integer) As Boolean
Function AggiornaConfigurazione(NomeTabella As String, NomeColonna As String, Utente As String, Contenuto As String, clausolaWhere As String, Linea As Integer) As Boolean

On Error GoTo Gesterrore

    'Dim Utente As String
    'Dim p As Integer

    'p = InStr(Contenuto, "|")
    'If p > 0 Then
        'Utente = Left(Contenuto, p - 1)
        'Contenuto = Mid(Contenuto, p + 1)
        'AggiornaConfigurazione = AggiornaCampoDB(Utente, NomeTabella, NomeColonna, Contenuto, clausolaWhere, Linea, Master_Slave)
        AggiornaConfigurazione = AggiornaCampoDB(Utente, NomeTabella, NomeColonna, Contenuto, clausolaWhere, Linea)
    'End If
Exit Function

Gesterrore:
    Call WindasLog("AggiornaConfigurazione " + Error(Err), 1, OPC)
    AggiornaConfigurazione = False

End Function

'luca aprile 2017
Function AggiornaCampoDB_old(Utente As String, NomeTabella As String, NomeColonna As String, Contenuto As String, clausolaWhere As String, Linea As Integer, Master_Slave As Integer) As Boolean

    Dim strSQL As String
    Dim OLDcontenuto As String
    
    'Alby Febbraio 2014 Gestione nuova tabella storicizzazione eventi
    Dim Chiave As String
    Dim Stazione As String
    Dim Data As String
    Dim Ora As String
    Dim Tipo As String
    Dim Descrizione As String
    Dim rsConfig As Object

    On Error GoTo Gesterrore
    
    Set rsConfig = CreateObject("AttimoFwk.CData")
    With rsConfig
        .SetDBType .Conn_SQL
        .SetServer CStr(AppServer)
        .SetDatabase CStr(AppDatabase), CStr(AppDBUser), CStr(AppDBPwd)
        .SetDbversion CStr(AppDbVersion)
        .SetMessages False
        .Err_Activate = True
    End With
    
    AggiornaCampoDB_old = True
    
    'luca 25/07/2016
    'Stazione = "WinCC"
    Stazione = CStr(Linea) & "_SiCEMS"
    Data = Format(Now, "yyyymmdd")
    Ora = Format(Now, "hh.nn.ss")
    
    If NomeTabella = "" Then
        'Azione pulsante
        Tipo = 0
        NomeTabella = "AZIONE"
        nomecampo = "PULSANTE"
        Descrizione = Contenuto
    Else
        'Modifica parametrizzazione
        Tipo = 1
    End If
    Chiave = Data + " " + Ora + " " + NomeTabella + " " + NomeColonna + " " + CStr(Master_Slave) + " " + Tipo
        
    'se si tratta di una modifica di parametrizzazione
    If Tipo = 1 Then
        strSQL = "SELECT " + NomeColonna + " FROM " + NomeTabella + " WHERE " + clausolaWhere
        rsConfig.selectionfast strSQL
            
        If Not rsConfig.iseof Then OLDcontenuto = rsConfig.getValue(NomeColonna)
            
        strSQL = "UPDATE " + NomeTabella + " SET " + NomeColonna + "= '" + Contenuto + "' WHERE " + clausolaWhere
        
        Call WindasLog("WindasFwk Query=" + strSQL, 0, OPC)
        rsConfig.ExecuteSql strSQL
    End If
    
    If UCase(NomeTabella) = "WAS_CONFIG" Then
        strSQL = "SELECT * FROM WAS_CONFIG WHERE " + clausolaWhere
        rsConfig.ExecuteSql strSQL
        If Not rsConfig.iseof Then Descrizione = rsConfig.getValue("cc_description")
    ElseIf UCase(NomeTabella) = "WAS_MEASURES" Then
        strSQL = "SELECT * FROM WAS_MEASURES WHERE " + clausolaWhere
        rsConfig.selectionfast strSQL
        If Not rsConfig.iseof Then Descrizione = rsConfig.getValue("c3")
    End If
    
    strSQL = "INSERT INTO wls_actionlog (Stazione,Chiave,Data,Ora,PC,Tipo,Descrizione,Tabella,Campo,ValorePrecedente,ValoreAttuale,Utente) VALUES ("
    strSQL = strSQL + "'" + Stazione + "','" + Chiave + "','" + Data + "','" + Ora + "'," + CStr(Master_Slave) + "," + Tipo + ",'" + Descrizione + "','" + NomeTabella + "','" + NomeColonna + "','"
    strSQL = strSQL + OLDcontenuto + "','" + Contenuto + "','" + Utente + "')"

    'Call WindasLog("WindasFwk Query=" + strSQL, 0, OPC)
    rsConfig.ExecuteSql strSQL
    
    Exit Function
    
Gesterrore:
    Call WindasLog("AggiornaCampoDB_old " + Error(Err), 1, OPC)
    AggiornaCampoDB_old = False


End Function
'luca aprile 2017 revisionata in seguito a gestione nuova tabella per inserimento modifica configurazione
Function AggiornaCampoDB(Utente As String, NomeTabella As String, NomeColonna As String, Contenuto As String, clausolaWhere As String, Linea As Integer) As Boolean

    Dim strSQL As String
    Dim OLDcontenuto As String
    
    Dim Stazione As String
    Dim Tipo As String
    Dim Descrizione As String
    Dim rsConfig As Object
    Dim codice As String
    Dim HeaderCampo As String
    Dim codice_wds_designer As String
    
    On Error GoTo Gesterrore
    
    NewDataObj rsConfig
    AggiornaCampoDB = True
    
    'luca 25/07/2016
    'Stazione = "WinCC"
    Stazione = CStr(Linea) & "CEMS1"
    
    If NomeTabella = "" Then
        'Azione pulsante
        Tipo = 0
        NomeTabella = "AZIONE"
        nomecampo = "PULSANTE"
        Descrizione = Contenuto
    Else
        'Modifica parametrizzazione
        Tipo = 1
    End If
        
    'se si tratta di una modifica di parametrizzazione
    If Tipo = 1 Then
        strSQL = "SELECT " + NomeColonna + " FROM " + NomeTabella + " WHERE " + clausolaWhere
        rsConfig.selectionfast strSQL
            
        If Not rsConfig.iseof Then OLDcontenuto = rsConfig.getValue(NomeColonna)
            
        strSQL = "UPDATE " + NomeTabella + " SET " + NomeColonna + "= '" + Contenuto + "' WHERE " + clausolaWhere
        
        Call WindasLog("WindasFwk Query=" + strSQL, 0, OPC)
        rsConfig.ExecuteSql strSQL
    End If
    
    If UCase(NomeTabella) = "WAS_CONFIG" Then
        strSQL = "SELECT * FROM WAS_CONFIG WHERE " + clausolaWhere
        rsConfig.ExecuteSql strSQL
        If Not rsConfig.iseof Then
            Descrizione = rsConfig.getValue("cc_description")
            codice = rsConfig.getValue("cc_code")
        End If
        codice_wds_designer = "GridEX1_G"
    ElseIf UCase(NomeTabella) = "WAS_MEASURES" Then
        strSQL = "SELECT * FROM WAS_MEASURES WHERE " + clausolaWhere
        rsConfig.selectionfast strSQL
        If Not rsConfig.iseof Then
            Descrizione = rsConfig.getValue("c3")
            codice = rsConfig.getValue("c2")
        End If
        codice_wds_designer = "GridEX1_S"
    End If
    
    strSQL = "SELECT * FROM WDS_DESIGNER WHERE DS_GROUP = '*' AND DS_CONTAINER = 'DlgStation' AND DS_PARENT = '" & codice_wds_designer & "' AND DS_OBJECT = '" & NomeColonna & "'"
    rsConfig.selectionfast strSQL
    If Not rsConfig.iseof Then
        HeaderCampo = Trim(rsConfig.getValue("DS_DESCRIPTION"))
    Else
        HeaderCampo = ""
    End If
    
    If OLDcontenuto <> Contenuto Then
        strSQL = "INSERT INTO WLS_CFGLOG (Station,Parameter,DescParameter, ColumnField, ColumnHeader,OldValue,NewValue,ActiveUser,Date,Time) VALUES "
        strSQL = strSQL + " ('" + Stazione + "','" + codice + "','" + Descrizione + "','" + NomeColonna + "','" + HeaderCampo + "','"
        strSQL = strSQL + OLDcontenuto + "','" + Contenuto + "','" + Utente + "','" + Format(Now, "yyyymmdd") + "','" + Format(Now, "hh.nn.ss") + "')"
        
        'Call WindasLog("WindasFwk Query=" + strSQL, 0, OPC)
        rsConfig.ExecuteSql strSQL
    End If
    
    Set rsConfig = Nothing
    
    Exit Function
    
Gesterrore:
    Call WindasLog("AggiornaCampoDB " + Error(Err), 1, OPC)
    AggiornaCampoDB = False


End Function
