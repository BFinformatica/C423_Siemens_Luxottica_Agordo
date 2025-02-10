Attribute VB_Name = "ArchiviazioneDB"
Option Explicit

Global TimeStamp As Date
Global NomeDBElementare As String
Global PathDBElementare As String
Global EtichettaTipoDato(10) As String

Sub SalvaDatiElementariDB(Optional ByVal InsertSingole As Boolean = False)
'Federica settembre 2017 - Passo un parametro per la chiamata a funzioni
'diverse per la costruzione della query di salvataggio dati
    
    On Error GoTo Gesterrore

    Dim rsConfig As Object
    NewDataObj rsConfig
    
    CRLF = Chr(13) + Chr(10)
    EtichettaTipoDato(0) = "M"
    EtichettaTipoDato(1) = "S"
        
    'luca giugno 2017
    Call SalvaDatiElementariDBgeneraMese(rsConfig)
    Call SalvaDatiElementariDBgeneraGiorno(rsConfig, 0)    'per dato misurato
    If InsertSingole Then
        Call SalvaDatiElementariDBarchiviaInsertSingole(rsConfig, 0)        'per dato misurato
    Else
        Call SalvaDatiElementariDBarchivia(rsConfig, 0)        'per dato misurato
    End If
    
    'luca giugno 2017
    Call GestioneBackupMensile
    Call GestioneBackupPrincipale
    
    Set rsConfig = Nothing
    
    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDB " + Error(Err), 1, OPC)
End Sub

'Federica settembre 2017 - Procedura per l'inserimento di dati recuperati da ADAM
Sub SalvaDatiElementariDBarchiviaInsertSingole(rsConfig, Tipo)

    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Dim iIdx As Integer
    Dim NomeTabella As String
    Dim sFormat As String
    Dim strQuery As String
        
    'Alby Aprile 2015 DB dati elementari
    NomeTabella = "BF" + EtichettaTipoDato(Tipo) + Format(TimeStamp, "yyyymmdd")
    
    'luca giugno 2017
    strSQL = ""
    strSQL = strSQL + "USE [" & NomeDBElementare & "]" + CRLF
    strSQL = strSQL + "BEGIN TRANSACTION" + CRLF
    strSQL = strSQL + "BEGIN TRY" + CRLF
    
    For iIdx = 0 To gnNroParametriStrumenti
        Select Case ParametriStrumenti(iIdx).NroDecimali
            Case 0
                sFormat = "0"
            Case 1
                sFormat = "0.0"
            Case 2
                sFormat = "0.00"
            Case 3
                sFormat = "0.000"
            Case Else
                sFormat = "0.0000"
        End Select
        
        If Not IsNull(ValIst(Tipo, iIdx)) And Not IsEmpty(ValIst(Tipo, iIdx)) Then
            
            strSQL = strSQL + "IF EXISTS("
            strSQL = strSQL + "SELECT * FROM " & NomeTabella & " "
            strSQL = strSQL + "WHERE DT_STATIONCODE = '" & gsClienteDi & "' "
            strSQL = strSQL + "AND DT_MEASURECOD = '" & ParametriStrumenti(iIdx).NomeParametro & "' "
            strSQL = strSQL + "AND DT_DATETIME = '" & Format(TimeStamp, "yyyymmddhhnnss") & "') "
            If Status(Tipo, iIdx) = "VAL" Then
                strSQL = strSQL + " UPDATE " & NomeTabella & " "
                strSQL = strSQL + "     SET DT_VALUE = " & Replace(Format(ValIst(Tipo, iIdx), sFormat), ",", ".") + ", "
                strSQL = strSQL + "         DT_VALIDFLAG = '" & Status(Tipo, iIdx) + "', "
                strSQL = strSQL + "         DT_VALUEN = " & Replace(Format(ValIst(1, iIdx), sFormat), ",", ".") + ", "
                strSQL = strSQL + "         DT_VALIDFLAGN = '" & Status(1, iIdx) + "' "
                strSQL = strSQL + "WHERE DT_STATIONCODE = '" & gsClienteDi & "' "
                strSQL = strSQL + "AND DT_MEASURECOD = '" & ParametriStrumenti(iIdx).NomeParametro & "' "
                strSQL = strSQL + "AND DT_DATETIME = '" & Format(TimeStamp, "yyyymmddhhnnss") & "' "
                strSQL = strSQL + "AND DT_VALIDFLAG = 'ERR' OR DT_VALIDFLAGN = 'ERR' "
            Else
                'Query fittizia altrimenti va in errore lo script
                strSQL = strSQL + " SELECT * FROM " & NomeTabella & " WHERE 0 = 1 "
            End If
            strSQL = strSQL + " ELSE "
            strSQL = strSQL + "INSERT INTO " & NomeTabella & "(DT_STATIONCODE, DT_MEASURECOD,DT_DATETIME,DT_VALUE,DT_VALIDFLAG,DT_VALUEN,DT_VALIDFLAGN,DATEHOUR) VALUES ('"
            strSQL = strSQL + gsClienteDi + "','"
            strSQL = strSQL + ParametriStrumenti(iIdx).NomeParametro + "','"
            strSQL = strSQL + Format(TimeStamp, "yyyymmddhhnnss") + "',"
            strSQL = strSQL + Replace(Format(ValIst(Tipo, iIdx), sFormat), ",", ".") + ",'"
            strSQL = strSQL + Status(Tipo, iIdx) + "'" + ","
            strSQL = strSQL + Replace(Format(ValIst(1, iIdx), sFormat), ",", ".") + ",'"
            strSQL = strSQL + Status(1, iIdx) + "'" + ",'"
            strSQL = strSQL + CreateDateForSQLS(TimeStamp) + "')" + CRLF
        End If
    Next iIdx
    
    strSQL = strSQL + "COMMIT TRANSACTION" + CRLF
    strSQL = strSQL + "END TRY" + CRLF
    strSQL = strSQL + "BEGIN CATCH" + CRLF
    strSQL = strSQL + "ROLLBACK TRANSACTION" + CRLF
    strSQL = strSQL + "END CATCH" + CRLF
    
    rsConfig.ExecuteSql strSQL
    
    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDBarchivia " + Error(Err), 1, OPC)
    Resume Next
    
End Sub

Sub SalvaDatiElementariDBgeneraAnno(rsConfig)
    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Dim FileSQL As String
        
    If Dir(PathDBElementare, vbDirectory) = "" Then
        MkDir PathDBElementare
    End If
    
    NomeDBElementare = gsClienteDi & "_" + Format(TimeStamp, "yyyy")
    
    FileSQL = PathDBElementare + "\" + NomeDBElementare
    
    strSQL = ""
    strSQL = strSQL + "USE [master]" + CRLF
    strSQL = strSQL + "CREATE DATABASE [" + NomeDBElementare + "] ON  PRIMARY " + CRLF
    strSQL = strSQL + "( NAME = '" + NomeDBElementare + "', FILENAME = '" + FileSQL + ".mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )" + CRLF
    strSQL = strSQL + "LOG ON " + CRLF
    strSQL = strSQL + "( NAME = '" + NomeDBElementare + "_LOG', FILENAME = '" + FileSQL + "_LOG.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)" + CRLF
    strSQL = strSQL + "COLLATE Latin1_General_CI_AS" + CRLF
    
    rsConfig.ExecuteSql strSQL

    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDBgeneraAnno " + Error(Err), 1, OPC)

End Sub

'luca giugno 2017
Sub SalvaDatiElementariDBgeneraMese(rsConfig)
    
    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Dim FileSQL As String
        
    If Dir(PathDBElementare, vbDirectory) = "" Then
        MkDir PathDBElementare
    End If
    
    NomeDBElementare = gsClienteDi & "_" + Format(TimeStamp, "yyyymm")
    
    FileSQL = PathDBElementare + "\" + NomeDBElementare
    FileSQL = Replace(FileSQL, "\\", "\")
    
    strSQL = ""
    strSQL = strSQL + "USE [master]" + CRLF
    strSQL = strSQL + "IF NOT EXISTS (SELECT name FROM master.sys.databases WHERE name = N'" & NomeDBElementare & "')" + CRLF
    strSQL = strSQL + "BEGIN" + CRLF
    strSQL = strSQL + "CREATE DATABASE [" + NomeDBElementare + "] ON  PRIMARY " + CRLF
    strSQL = strSQL + "( NAME = '" + NomeDBElementare + "', FILENAME = '" + FileSQL + ".mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )"
    strSQL = strSQL + "LOG ON " + CRLF
    strSQL = strSQL + "( NAME = N'" & NomeDBElementare & "_log', FILENAME = N'" + FileSQL + ".ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)" + CRLF
    strSQL = strSQL + "COLLATE Latin1_General_CI_AS" + CRLF
    strSQL = strSQL + "ALTER DATABASE [" + NomeDBElementare + "] SET AUTO_CLOSE OFF" + CRLF
    strSQL = strSQL + "END" + CRLF
    
    rsConfig.ExecuteSql strSQL

    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDBgeneraMese " + Error(Err), 1, OPC)

End Sub

Sub SalvaDatiElementariDBarchivia(rsConfig, Tipo)
    
    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Dim iIdx As Integer
    Dim NomeTabella As String
    Dim sFormat As String
        
    'Alby Aprile 2015 DB dati elementari
    NomeTabella = "BF" + EtichettaTipoDato(Tipo) + Format(TimeStamp, "yyyymmdd")
    
    'luca giugno 2017
    strSQL = ""
    strSQL = strSQL + "USE [" & NomeDBElementare & "]" + CRLF
    strSQL = strSQL + "BEGIN TRANSACTION" + CRLF
    strSQL = strSQL + "BEGIN TRY" + CRLF
    
    For iIdx = 0 To gnNroParametriStrumenti
        Select Case ParametriStrumenti(iIdx).NroDecimali
            Case 0
                sFormat = "0"
            Case 1
                sFormat = "0.0"
            Case 2
                sFormat = "0.00"
            Case 3
                sFormat = "0.000"
            Case Else
                sFormat = "0.0000"
        End Select
        
        'luca giugno 2017
        If Not IsNull(ValIst(Tipo, iIdx)) And Not IsEmpty(ValIst(Tipo, iIdx)) Then
            'Federica Novembre 2018 - Salvataggio elemetari normalizzati
'            strSQL = strSQL + "INSERT INTO " & NomeTabella & "(DT_STATIONCODE, DT_MEASURECOD,DT_DATETIME,DT_VALUE,DT_VALIDFLAG,DATEHOUR) VALUES ('"
'            strSQL = strSQL + gsClienteDi + "','"
'            strSQL = strSQL + ParametriStrumenti(iIdx).NomeParametro + "','"
'            strSQL = strSQL + Format(TimeStamp, "yyyymmddhhnnss") + "',"
'            strSQL = strSQL + Replace(Format(ValIst(Tipo, iIdx), sFormat), ",", ".") + ",'"
'            strSQL = strSQL + Status(Tipo, iIdx) + "'" + ",'"
'            strSQL = strSQL + CreateDateForSQLS(TimeStamp) + "')" + CRLF
            strSQL = strSQL + "INSERT INTO " & NomeTabella & "(DT_STATIONCODE, DT_MEASURECOD, DT_DATETIME, DT_VALUE, DT_VALIDFLAG, DT_VALUEN, DT_VALIDFLAGN, DATEHOUR) VALUES ('"
            strSQL = strSQL + gsClienteDi + "','"
            strSQL = strSQL + ParametriStrumenti(iIdx).NomeParametro + "','"
            strSQL = strSQL + Format(TimeStamp, "yyyymmddhhnnss") + "',"
            strSQL = strSQL + Replace(Format(ValIst(Tipo, iIdx), sFormat), ",", ".") + ",'"
            strSQL = strSQL + Status(Tipo, iIdx) + "'" + ","
            'Federica novembre 2018 - Salvataggio elementari normalizzati
            strSQL = strSQL + Replace(Format(ValIst(1, iIdx), sFormat), ",", ".") + ",'"
            strSQL = strSQL + Status(1, iIdx) + "'" + ",'"
            '-------
            strSQL = strSQL + CreateDateForSQLS(TimeStamp) + "')" + CRLF

        End If
        
    Next iIdx
    
    strSQL = strSQL + "COMMIT TRANSACTION" + CRLF
    strSQL = strSQL + "END TRY" + CRLF
    strSQL = strSQL + "BEGIN CATCH" + CRLF
    strSQL = strSQL + "ROLLBACK TRANSACTION" + CRLF
    strSQL = strSQL + "END CATCH" + CRLF
    
    rsConfig.ExecuteSql strSQL
    
    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDBarchivia " + Error(Err), 1, OPC)
    Resume Next

End Sub

'luca giugno 2017
Sub GestioneBackupMensile()

    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Static backupMonthDone As Boolean
        
    If hour(TimeStamp) = NumeroLinea And minute(TimeStamp) = 15 Then
        If Not backupMonthDone Then
            If Dir("C:\Windas\BackupDB", vbDirectory) = "" Then
                MkDir "C:\Windas\BackupDB"
            End If
            'database del giorno precedente, per salvare a capodanno quello dell'anno prima
            NomeDBElementare = gsClienteDi & "_" + Format(DateAdd("d", -1, TimeStamp), "yyyymm")
            strSQL = "use [" & NomeDBElementare & "]; backup database [" & NomeDBElementare & "] to disk = 'C:\Windas\BackupDB\" & NomeDBElementare & ".bak' WITH INIT;"
            
            Shell ("osql -S " & connDB(iConnDBDefault).AppServer & " -U " & connDB(iConnDBDefault).AppDBUser & " -P " & connDB(iConnDBDefault).AppDBPwd & " -d -Q """ & strSQL & """")
            Call WindasLog("Lanciato backup dati elementari db:" + NomeDBElementare, 0, OPC)
            backupMonthDone = True
        End If
    Else
        backupMonthDone = False
    End If
    
    Exit Sub
Gesterrore:
    Call WindasLog("GestioneBackupMensile " + Error(Err), 1, OPC)

End Sub

'luca giugno 2017
Sub GestioneBackupPrincipale()

    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Static backupPrincipalDone As Boolean
        
    If hour(TimeStamp) = 0 And minute(TimeStamp) = NumeroLinea * 5 Then
        If Not backupPrincipalDone Then
            If Dir("C:\Windas\BackupDB", vbDirectory) = "" Then
                MkDir "C:\Windas\BackupDB"
            End If
            'database del giorno precedente, per salvare a capodanno quello dell'anno prima
            NomeDBElementare = connDB(iConnDBDefault).AppDatabase
            strSQL = "use [" & NomeDBElementare & "]; backup database [" & NomeDBElementare & "] to disk = 'C:\Windas\BackupDB\" & NomeDBElementare & ".bak' WITH INIT;"
            
            Shell ("osql -S " & connDB(iConnDBDefault).AppServer & " -U " & connDB(iConnDBDefault).AppDBUser & " -P " & connDB(iConnDBDefault).AppDBPwd & " -d -Q """ & strSQL & """")
            Call WindasLog("Lanciato backup principale db:" + NomeDBElementare, 0, OPC)
            backupPrincipalDone = True
        End If
    Else
        backupPrincipalDone = False
    End If
    
    Exit Sub
Gesterrore:
    Call WindasLog("GestioneBackupPrincipale " + Error(Err), 1, OPC)

End Sub

Sub SalvaDatiElementariDBgeneraGiorno(rsConfig, Tipo)
    On Error GoTo Gesterrore
    
    Dim strSQL As String
    Dim NomeTabella As String
    
    NomeTabella = "BF" + EtichettaTipoDato(Tipo) + Format(TimeStamp, "yyyymmdd")
    
    'luca giugno 2017
    strSQL = ""
    strSQL = strSQL + "IF NOT (EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & NomeDBElementare & "' AND  TABLE_NAME = '" & NomeTabella & "'))" + CRLF
    strSQL = strSQL + "BEGIN" + CRLF
    strSQL = strSQL + "USE [" & NomeDBElementare & "]" + CRLF
    strSQL = strSQL + "SET ANSI_NULLS ON" + CRLF
    strSQL = strSQL + "SET QUOTED_IDENTIFIER ON" + CRLF
    strSQL = strSQL + "SET ANSI_PADDING ON" + CRLF
    strSQL = strSQL + "CREATE TABLE [dbo].[" + NomeTabella + "](" + CRLF
    strSQL = strSQL + "[Id] [uniqueidentifier] NOT NULL," + CRLF 'Nicolò Novembre 2016
    strSQL = strSQL + "[DT_STATIONCODE] [varchar](10) COLLATE Latin1_General_CI_AS NOT NULL," + CRLF
    strSQL = strSQL + "[DT_MEASURECOD] [varchar](10) COLLATE Latin1_General_CI_AS NOT NULL," + CRLF
    strSQL = strSQL + "[DT_DATETIME] [varchar](14) COLLATE Latin1_General_CI_AS NOT NULL," + CRLF
    strSQL = strSQL + "[DT_VALUE] [float] NULL," + CRLF
    strSQL = strSQL + "[DT_VALIDFLAG] [varchar](50) COLLATE Latin1_General_CI_AS NULL," + CRLF
    'Federica Novembre 2018 - Aggiunta campi per salvataggio dati Elementari Normalizzati
    strSQL = strSQL + "[DT_VALUEN] [float] NULL," + CRLF
    strSQL = strSQL + "[DT_VALIDFLAGN] [varchar](50) COLLATE Latin1_General_CI_AS NULL," + CRLF
    '--------------------
    strSQL = strSQL + "[DateHour] [datetime] NULL," + CRLF ' Nicolò Novembre 2016
    strSQL = strSQL + "CONSTRAINT [PK_DAT_" + NomeTabella + "] PRIMARY KEY CLUSTERED " + CRLF
    strSQL = strSQL + "(" + CRLF
    strSQL = strSQL + "[Id] ASC" + CRLF
    strSQL = strSQL + ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" + CRLF
    strSQL = strSQL + ") ON [PRIMARY]" + CRLF
    strSQL = strSQL + "SET ANSI_PADDING OFF" + CRLF
    strSQL = strSQL + "CREATE UNIQUE NONCLUSTERED INDEX [PRIMARY] ON [dbo].[" + NomeTabella + "]" + CRLF
    strSQL = strSQL + "(" + CRLF
    strSQL = strSQL + "[DT_STATIONCODE] ASC," + CRLF
    strSQL = strSQL + "[DT_MEASURECOD] ASC," + CRLF
    strSQL = strSQL + "[DT_DATETIME] ASC" + CRLF
    strSQL = strSQL + ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" + CRLF
    strSQL = strSQL + "ALTER TABLE [dbo].[" + NomeTabella + "] ADD  CONSTRAINT [DF_dbo." + NomeTabella + "_Id]  DEFAULT (newsequentialid()) FOR [Id]" + CRLF
    strSQL = strSQL + "END" + CRLF
    
    rsConfig.ExecuteSql strSQL

    Exit Sub
Gesterrore:
    Call WindasLog("SalvaDatiElementariDBgeneraGiorno " + Error(Err), 1, OPC)

End Sub
