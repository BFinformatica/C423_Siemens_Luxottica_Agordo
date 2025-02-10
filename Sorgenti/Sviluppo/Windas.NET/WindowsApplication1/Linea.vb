Class Linea
    Public CodiceLinea As String
    Public NumeroLinea As String
    Public NomeLinea As String
    Public Misure() As Misura
    Public NrMisure As Integer
    Public Allarmi(0, 0) As String
    Public NrAllarmi As Integer
    Public Stati(0, 0) As String
    Public NrStati As Integer
    Public Soglie(0, 0) As String
    Public NrSoglie As Integer
    Public Calibrazioni(0) As String
    Public NrCalibrazioni As Integer

    Public tmpSoglie(3) As Campo
    Public tmpQAL2(4) As Campo

    Public Sub New()

        tmpSoglie(0).Campo = "C11"
        tmpSoglie(0).Chiave = "SogliaAttenzione"
        tmpSoglie(1).Campo = "C12"
        tmpSoglie(1).Chiave = "SogliaAllarme"
        tmpSoglie(2).Campo = "C75"
        tmpSoglie(2).Chiave = "SogliaAttenzioneGiorno"
        tmpSoglie(3).Campo = "C76"
        tmpSoglie(3).Chiave = "SogliaAllarmeGiorno"

        tmpQAL2(0).Campo = "C41"
        tmpQAL2(0).Chiave = "M"
        tmpQAL2(1).Campo = "C42"
        tmpQAL2(1).Chiave = "Q"
        tmpQAL2(2).Campo = "C43"
        tmpQAL2(2).Chiave = "Range"
        tmpQAL2(3).Campo = "C44"
        tmpQAL2(3).Chiave = "IC"
        tmpQAL2(4).Campo = "C45"
        tmpQAL2(4).Chiave = "DataQAL2"

    End Sub

    Public Sub CaricaMisure()

        Dim rsDati As DataTable
        Dim i As Integer
        Dim nr As Integer
        Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

        Try
            'Estrazione dati misure
            rsDati = DB.Reader("SELECT was_measures.*,wds_gentab.gt_description FROM was_measures inner join wds_gentab on c2=gt_code where gt_type = 'params' AND cm_stationcode = '" & CodiceLinea & "' order by C1")

            NrMisure = rsDati.Rows.Count - 1
            i = 0
            ReDim Misure(NrMisure)
            For Each r In rsDati.Rows
                With Misure(i)

                    ReDim .Soglie(tmpSoglie.Length - 1)
                    tmpSoglie.CopyTo(.Soglie, 0)
                    ReDim .QAL2(tmpQAL2.Length - 1)
                    tmpQAL2.CopyTo(.QAL2, 0)

                    .Indice = r("C1")
                    .Codice = r("C2")
                    .FondoScala = r("c10")
                    .Descrizione = Trim(r("gt_description")) + " (" + Trim(r("c4")) + ")"
                    For nr = 0 To UBound(.Soglie)
                        .Soglie(nr).Valore = CStr(r(.Soglie(nr).Campo))
                    Next

                    .ZeroRif = CStr(r("C58"))
                    .SpanRif = CStr(r("C59"))

                    For nr = 0 To UBound(.QAL2)
                        .QAL2(nr).Valore = IIf(IsDBNull(r(.QAL2(nr).Campo)), "", r(.QAL2(nr).Campo))
                        If (.QAL2(nr).Chiave = "DataQAL2") Then
                            If (.QAL2(nr).Valore = String.Empty) Then
                                .QAL2(nr).Valore = Now
                            Else
                                .QAL2(nr).Valore = Date.ParseExact(.QAL2(nr).Valore, "yyyymmdd", provider)
                            End If
                        Else
                            If (.QAL2(nr).Valore = String.Empty) Then .QAL2(nr).Valore = "0"
                        End If
                    Next

                    i = i + 1
                End With
            Next
            rsDati = Nothing

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Linea.CaricaMisure errore: " + ex.ToString)
        End Try

    End Sub

    Public Sub CaricaAllarmi()

        Dim rsDati As DataTable
        Dim i As Integer

        Try
            'Estrazione allarmi
            rsDati = DB.Reader("SELECT was_digital.*, wds_gentab.gt_description, wds_gentab.gt_str5 " & _
             "FROM was_digital inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'alarm' AND cd_stationcode = '" & Me.CodiceLinea & "' AND was_digital.C13 = '1' order by C1")

            NrAllarmi = rsDati.Rows.Count - 1
            ReDim Allarmi(NrAllarmi, 1)
            i = 0
            For Each r In rsDati.Rows
                Allarmi(i, 0) = r("gt_description")
                Allarmi(i, 1) = r("C1")

                i = i + 1
            Next
            rsDati = Nothing

            'Estrazione stati
            rsDati = DB.Reader("SELECT was_digital.*, wds_gentab.gt_description, wds_gentab.gt_str5 " & _
             "FROM was_digital inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'alarm' AND cd_stationcode = '" & Me.CodiceLinea & "' AND was_digital.C13 = '2' order by C1")

            NrStati = rsDati.Rows.Count - 1
            ReDim Stati(NrStati, 1)
            i = 0
            For Each r In rsDati.Rows
                Stati(i, 0) = r("gt_description")
                Stati(i, 1) = r("C1")

                i = i + 1
            Next
            rsDati = Nothing

            'Estrazione soglie
            rsDati = DB.Reader("SELECT was_digital.*, wds_gentab.gt_description, wds_gentab.gt_str5 " & _
             "FROM was_digital inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'alarm' AND cd_stationcode = '" & Me.CodiceLinea & "' AND was_digital.C13 = '3' order by C1")

            NrSoglie = rsDati.Rows.Count - 1
            ReDim Soglie(NrSoglie, 1)
            i = 0
            For Each r In rsDati.Rows
                Soglie(i, 0) = r("gt_description")
                Soglie(i, 1) = r("C1")

                i = i + 1
            Next
            rsDati = Nothing

            'Estrazione calibrazione
            rsDati = DB.Reader("SELECT was_digital.*, wds_gentab.gt_description, wds_gentab.gt_str5 " & _
             "FROM was_digital inner join wds_gentab on c2=gt_code " & _
             "where gt_type = 'alarm' AND cd_stationcode = '" & Me.CodiceLinea & "' AND was_digital.C13 = '4' order by C1")

            NrCalibrazioni = rsDati.Rows.Count - 1
            ReDim Calibrazioni(NrCalibrazioni)
            i = 0
            For Each r In rsDati.Rows
                Calibrazioni(i) = r("C1")

                i = i + 1
            Next
            rsDati = Nothing

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Linea.CaricaAllarmi errore: " + ex.ToString)
        End Try

    End Sub

    Public Sub ConfigurazioneAggiornaDati(ByVal Indice_misura As Integer)

        Dim conn As GestioneDatabase.TipologiaConnessione
        Dim InErrore As Boolean
        Dim PartnerOK As Boolean

        Try
            PartnerOK = CBool(LeggiTag("ANOMALIA_PARTNER"))

            'Prima va fatto sul pc remoto.
            If PartnerOK Then
                For Each conn In GestDB.Connections
                    If Not conn.IsDefault Then
                        'Aggiorno QAL2
                        If Not SoglieAggiornaDati(Indice_misura, conn) Then
                            Call WindasLog(MSG_LOG, "ConfigurazioneAggiornaDati: query aggiornamento configurazione su " & conn.Server & " non eseguita correttamente")
                            InErrore = True
                            Exit For
                        End If
                    End If
                Next
            End If

            If Not InErrore Then
                If Not SoglieAggiornaDati(Indice_misura, DB.Connessione) Then
                    Call WindasLog(MSG_LOG, "ConfigurazioneAggiornaDati: query aggiornamento configurazione su " & DB.Connessione.Server & " non eseguita correttamente")
                End If
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "ConfigurazioneAggiornaDati errore: " + ex.ToString)

        End Try

    End Sub

    Private Function SoglieAggiornaDati(Indice_Misura As Integer, ByVal conn As GestioneDatabase.TipologiaConnessione) As Boolean

        Dim strSQL As String
        Dim cmpSoglia As Campo

        SoglieAggiornaDati = True
        Try
            Dim tmpDB As GestioneDatabase.Database = GestioneDatabase.TipologiaConnessione.GetDatabase(conn)

            'Faccio una query UNICA di aggiornamento
            strSQL = "UPDATE was_measures SET "
            For iCampo = 0 To UBound(tmpSoglie)
                strSQL = strSQL & tmpSoglie(iCampo).Campo & " = '" & tmpSoglie(iCampo).Valore & IIf(iCampo <> UBound(tmpSoglie), "', ", "' ")
            Next
            strSQL = strSQL + "WHERE C2 = '" & Misure(Indice_Misura).Codice & "' AND CM_StationCode = '" & CodiceLinea & "'"
            If tmpDB.TryQuery(strSQL) > 0 Then
                'Se la query è andata a buon fine, salvo un riga per parametro nel log
                For Each cmp In tmpSoglie
                    cmpSoglia = Misure(Indice_Misura).GetSogliaByChiave(cmp.Chiave)
                    If cmpSoglia.Valore <> cmp.Valore Then
                        Call AggiornaLogCambiamenti(tmpDB, "was_measures", Misure(Indice_Misura).Codice, Misure(Indice_Misura).Descrizione, cmp.Campo, cmpSoglia.Valore, cmp.Valore)
                    End If
                Next
            Else
                Call WindasLog(MSG_LOG, "Linea.SoglieAggiornaDati: query aggiornamento Soglie per il parametro " & Misure(Indice_Misura).Codice & " non eseguita correttamente")
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Linea.SoglieAggiornaDati errore: " + ex.ToString)
            SoglieAggiornaDati = False

        End Try

    End Function

    Private Function QAL2AggiornaDati(Indice_Misura As Integer, ByVal conn As GestioneDatabase.TipologiaConnessione) As Boolean

        Dim strSQL As String
        Dim cmpQAL As Campo
        Dim DBcampo As String

        QAL2AggiornaDati = True
        Try
            Dim tmpDB As GestioneDatabase.Database = GestioneDatabase.TipologiaConnessione.GetDatabase(conn)

            'Faccio una query UNICA di aggiornamento
            strSQL = "UPDATE was_measures SET "
            For iCampo = 0 To UBound(tmpQAL2)
                strSQL = strSQL & tmpQAL2(iCampo).Campo & " = '" & tmpQAL2(iCampo).Valore & IIf(iCampo <> UBound(tmpQAL2), "', ", "' ")
            Next
            strSQL = strSQL + "WHERE C2 = '" & Misure(Indice_Misura).Codice & "' AND CM_StationCode = '" & CodiceLinea & "'"
            If tmpDB.TryQuery(strSQL) > 0 Then
                'Se la query è andata a buon fine, salvo un riga per parametro nel log
                For Each cmp In tmpQAL2
                    cmpQAL = Misure(Indice_Misura).GetQAL2ByChiave(cmp.Chiave)
                    If cmpQAL.Chiave = "DataQAL2" Then
                        DBcampo = Format(cmpQAL.Valore, "yyyymmdd")
                    Else
                        DBcampo = cmpQAL.Valore
                    End If
                    If DBcampo <> cmp.Valore Then
                        Call AggiornaLogCambiamenti(tmpDB, "was_measures", Misure(Indice_Misura).Codice, Misure(Indice_Misura).Descrizione, cmp.Campo, DBcampo, cmp.Valore)
                    End If
                Next
            Else
                Call WindasLog(MSG_LOG, "Linea.QAL2AggiornaDati: query aggiornamento QAL2 per il parametro " & Misure(Indice_Misura).Codice & " non eseguita correttamente")
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Linea.QAL2AggiornaDati errore: " + ex.ToString)
            QAL2AggiornaDati = False

        End Try

    End Function

    Public Sub QAL2QAL3AggiornaDati(Indice_misura As Integer)

        Dim queryOK(6) As Boolean
        Dim InErrore As Boolean
        Dim PartnerOK As Boolean

        Try
            PartnerOK = CBool(LeggiTag("ANOMALIA_PARTNER"))

            '**** QAL2 ****
            'Prima va fatto sul pc remoto.
            If PartnerOK Then
                For Each conn In GestDB.Connections
                    If Not conn.IsDefault Then
                        'Aggiorno QAL2
                        If Not QAL2AggiornaDati(Indice_misura, conn) Then
                            Call WindasLog(MSG_LOG, "ConfigurazioneAggiornaDati: query aggiornamento configurazione su " & conn.Server & " non eseguita correttamente")
                            InErrore = True
                            Exit For
                        End If
                    End If
                Next
            End If

            If Not InErrore Then
                If Not QAL2AggiornaDati(Indice_misura, DB.Connessione) Then
                    Call WindasLog(MSG_LOG, "ConfigurazioneAggiornaDati: query aggiornamento configurazione su " & DB.Connessione.Server & " non eseguita correttamente")
                End If
            End If

            '**** QAL3 ****
            'queryOK(5) = AggiornaConfigurazione(CodiceLinea, CodiceLinea, "Windas.NET", "WAS_MEASURES", "C58", CStr(zeroRef), "C2='" & misura.Codice & "' AND CM_StationCode = '" & CodiceLinea & "'")
            'queryOK(6) = AggiornaConfigurazione(CodiceLinea, CodiceLinea, "Windas.NET", "WAS_MEASURES", "C59", CStr(spanRef), "C2='" & misura.Codice & "' AND CM_StationCode = '" & CodiceLinea & "'")

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Linea.QAL2QAL3AggiornaDati errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub AggiornaLogCambiamenti(ByVal database As GestioneDatabase.Database, ByVal nomeTabella As String, ByVal codice_parametro As String, descrizione_parametro As String, ByVal nome_campo As String, ByVal old_valore As String, ByVal new_valore As String)

        Dim HeaderCampo As String
        Dim strSQL As String
        Dim codice_wds_designer As String = ""
        Dim tmpTBL As DataTable

        Try
            Select Case UCase(nomeTabella)
                Case "WAS_CONFIG"
                    codice_wds_designer = "GridEX1_G"

                Case "WAS_MEASURES"
                    codice_wds_designer = "GridEX1_S"
            End Select

            tmpTBL = database.Reader("SELECT * FROM WDS_DESIGNER WHERE DS_GROUP = '*' AND DS_CONTAINER = 'DlgStation' AND DS_PARENT = '" & codice_wds_designer & "' AND DS_OBJECT = '" & nome_campo & "'")
            If tmpTBL.Rows.Count <> 0 Then
                HeaderCampo = Trim(tmpTBL(0)("DS_DESCRIPTION"))
            Else
                HeaderCampo = ""
            End If

            strSQL = "INSERT INTO WLS_CFGLOG (Station,Parameter,DescParameter, ColumnField, ColumnHeader,OldValue,NewValue,ActiveUser,Date,Time) VALUES "
            strSQL = strSQL + " ('" + CodiceLinea + "','" + codice_parametro + "','" + descrizione_parametro + "','" + nome_campo + "','" + HeaderCampo + "','"
            strSQL = strSQL + old_valore + "','" + new_valore + "','Windas.NET','" + Format(Now, "yyyyMMdd") + "','" + Format(Now, "hh.mm.ss") + "')"
            database.Query(strSQL)
        Catch ex As Exception
            Call WindasLog(ERR_LOG, "AggiornaLogCambiamenti errore: " & ex.ToString)

        End Try

    End Sub

End Class
