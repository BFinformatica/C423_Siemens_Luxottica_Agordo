
Public Class FrmMain

    Structure OLDLinee
        Dim PosizioneLinea As Integer
        Dim CodiceLinea As String
        Dim NomeLinea As String
    End Structure

    Dim gLineaSelzionata As Integer
    Dim arrayMisureQAL2QAL3(2) As String
    Dim Linea(50) As OLDLinee
    Dim ColoreON(128) As Double
    Dim ColoreOFF(128) As Double
    Dim FondoScala(128) As Double

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'CefSharp.Cef.Initialize(New WinForms.CefSettings());

        'Alby Marzo 2018
        Call WindasLog(MSG_LOG, "Windas.NET avviato")
        Call InizializzaComunicator()

        'Acquisisco connessioni
        Call GetConnectionParams()

        'luca marzo 2018
        Call GestioneLinee()

        'Federica settembre 2018 - L'impianto parte sempre come FERMO
        Call CambiaStatoImpianto(StatiImpianto.IMPIANTO_FERMO)

        'Alby Marzo 2018
        Call CaricaDati()

    End Sub

    Public Sub VisualizzaDati_TabMisure()

        Dim iIndice As Integer
        Dim ctrl As Control

        Try
            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrMisure
                    'Inserimento dati in Pagina Misure
                    ctrl = GetControlByName("Label" + .Misure(iIndice).Indice, TabMisure)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).Descrizione
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabMisure errore: " + ex.ToString)

        End Try

    End Sub

    Public Sub VisualizzaDati_TabQAL2QAL3()

        Dim iIndice As Integer
        Dim ctrl

        Try
            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrMisure
                    'Inserimento dati in Pagina QAL2/QAL3
                    ctrl = GetControlByName("txtM" & .Misure(iIndice).Indice, TabQAL)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetQAL2ByChiave("M").Valore

                    ctrl = GetControlByName("txtQ" & .Misure(iIndice).Indice, TabQAL)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetQAL2ByChiave("Q").Valore

                    ctrl = GetControlByName("txtIC" & .Misure(iIndice).Indice, TabQAL)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetQAL2ByChiave("IC").Valore

                    ctrl = GetControlByName("txtRange" & .Misure(iIndice).Indice, TabQAL)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetQAL2ByChiave("Range").Valore

                    ctrl = GetControlByName("dtpDataQAL2" & .Misure(iIndice).Indice, TabQAL)
                    If Not ctrl Is Nothing Then ctrl.Value = .Misure(iIndice).GetQAL2ByChiave("DataQAL2").Valore

                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabQAL2QAL3 errore: " + ex.ToString)

        End Try

    End Sub

    Sub VisualizzaDati_TabConfigurazione()

        Dim iIndice As Integer
        Dim ctrl As Control

        Try
            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrMisure
                    'Inserimento dati in Pagina Configurazione
                    ctrl = GetControlByName("lblConf" + .Misure(iIndice).Indice, TabConfigurazione)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).Descrizione

                    ctrl = GetControlByName("txtSogliaAtt_" & CStr(iIndice), TabConfigurazione)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetSogliaByChiave("SogliaAttenzione").Valore

                    ctrl = GetControlByName("txtSogliaAll_" & CStr(iIndice), TabConfigurazione)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetSogliaByChiave("SogliaAllarme").Valore

                    ctrl = GetControlByName("txtSogliaAttGrn_" & CStr(iIndice), TabConfigurazione)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetSogliaByChiave("SogliaAttenzioneGiorno").Valore

                    ctrl = GetControlByName("txtSogliaAllGrn_" & CStr(iIndice), TabConfigurazione)
                    If Not ctrl Is Nothing Then ctrl.Text = .Misure(iIndice).GetSogliaByChiave("SogliaAllarmeGiorno").Valore
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabConfigurazione errore: " + ex.ToString)

        End Try

    End Sub

    Sub VisualizzaDati_TabSinottico()

        Dim iIndice As Integer
        Dim ctrl As Control

        Try
            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrMisure
                    'Inserimento dati in Pagina Sinottico
                    ctrl = GetControlByName("lblSinottico" + .Misure(iIndice).Indice, TabSinottico)
                    If Not ctrl Is Nothing Then
                        ctrl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                        ctrl.Text = .Misure(iIndice).Descrizione
                    End If
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabSinottico errore: " + ex.ToString)

        End Try

    End Sub

    Sub VisualizzaDati_TabAllarmi()

        Dim iIndice As Integer
        Dim L As New Label
        Dim Indirizzo As Integer
        Dim Colonna As Integer
        Dim iIdx As Integer

        Try
            'Azzero
            TabAllarmi.Controls.Clear()

            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrAllarmi
                    L = New Label
                    TabAllarmi.Controls.Add(L)
                    L.Left = 80 + (600 * Colonna)
                    L.Top = 30 + (iIdx * 40)
                    L.Width = 560
                    L.Height = 35
                    L.BorderStyle = 2
                    L.TextAlign = ContentAlignment.MiddleCenter
                    L.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    L.Text = .Allarmi(iIndice, 0)
                    L.ForeColor = Color.White
                    L.BackColor = Color.DarkGreen
                    'L.Name = "D" + Format(Indirizzo, "000")
                    L.Name = "D" + .Allarmi(iIndice, 1)
                    AddHandler L.Click, AddressOf Label_click

                    iIdx += 1
                    Indirizzo += 1
                    If iIdx > 18 Then
                        If Colonna > 0 Then
                            iIdx = 0 : Colonna = 2
                        Else
                            iIdx = 0 : Colonna = 1
                        End If
                    End If
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabAllarmi errore: " + ex.ToString)

        End Try

    End Sub

    Sub VisualizzaDati_TabStati()

        Dim iIndice As Integer
        Dim L As New Label
        Dim Indirizzo As Integer
        Dim Colonna As Integer
        Dim iIdx As Integer

        Try
            'Azzero
            TabStati.Controls.Clear()

            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrStati
                    L = New Label
                    TabStati.Controls.Add(L)
                    L.Left = 80 + (600 * Colonna)
                    L.Top = 30 + (iIdx * 40)
                    L.Width = 560
                    L.Height = 35
                    L.BorderStyle = 2
                    L.TextAlign = ContentAlignment.MiddleCenter
                    L.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    L.Text = .Stati(iIndice, 0)
                    L.ForeColor = Color.White
                    L.BackColor = Color.DarkGreen
                    'L.Name = "D" + Format(Indirizzo, "000")
                    L.Name = "D" + .Stati(iIndice, 1)
                    AddHandler L.Click, AddressOf Label_click

                    iIdx += 1
                    Indirizzo += 1
                    If iIdx > 18 Then
                        If Colonna > 0 Then
                            iIdx = 0 : Colonna = 2
                        Else
                            iIdx = 0 : Colonna = 1
                        End If
                    End If
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabStati errore: " + ex.ToString)

        End Try

    End Sub

    Sub VisualizzaDati_TabSoglie()

        Dim iIndice As Integer
        Dim L As New Label
        Dim Indirizzo As Integer
        Dim Colonna As Integer
        Dim iIdx As Integer

        Try
            'Azzero
            TabSoglie.Controls.Clear()

            With Linee(gLineaSelzionata)
                For iIndice = 0 To .NrSoglie
                    L = New Label
                    TabSoglie.Controls.Add(L)
                    L.Left = 80 + (600 * Colonna)
                    L.Top = 30 + (iIdx * 40)
                    L.Width = 560
                    L.Height = 35
                    L.BorderStyle = 2
                    L.TextAlign = ContentAlignment.MiddleCenter
                    L.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    L.Text = .Soglie(iIndice, 0)
                    L.ForeColor = Color.White
                    L.BackColor = Color.DarkGreen
                    'L.Name = "D" + Format(Indirizzo, "000")
                    L.Name = "D" + .Soglie(iIndice, 1)
                    AddHandler L.Click, AddressOf Label_click

                    iIdx += 1
                    Indirizzo += 1
                    If iIdx > 18 Then
                        If Colonna > 0 Then
                            iIdx = 0 : Colonna = 2
                        Else
                            iIdx = 0 : Colonna = 1
                        End If
                    End If
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "VisualizzaDati_TabSoglie errore: " + ex.ToString)

        End Try

    End Sub

    Sub CaricaDati()

        Dim iIdx As Integer
        Dim ctrl As Control

        Try
            'Carico i dati per ciascuna linea
            For iIdx = 0 To UBound(Linee)
                Linee(iIdx).CaricaMisure()
                Linee(iIdx).CaricaAllarmi()

                'Nome linea a video
                ctrl = GetControlByName("lblLinea" & Linee(iIdx).NumeroLinea, Me)
                If Not ctrl Is Nothing Then ctrl.Text = Linee(iIdx).NomeLinea
            Next

            Call VisualizzaDati_TabSinottico()
            Call VisualizzaDati_TabMisure()
            Call VisualizzaDati_TabAllarmi()
            Call VisualizzaDati_TabStati()
            Call VisualizzaDati_TabSoglie()
            Call VisualizzaDati_TabQAL2QAL3()
            Call VisualizzaDati_TabConfigurazione()

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "CaricaDati errore: " + ex.ToString)

        End Try

    End Sub

    'luca marzo 2018
    Sub GestioneLinee()

        Dim rsDati As DataTable
        Dim numeroLinee As Integer
        Dim i As Integer = 0

        Try
            'Conto quante linee ci sono
            rsDati = DB.Reader("SELECT COUNT(*) as NumeroLinee FROM WDS_GENTAB WHERE GT_TYPE = 'STATIONS'")
            numeroLinee = CInt(rsDati.Rows(0)("NumeroLinee"))

            If numeroLinee = 0 Then
                Call WindasLog(ERR_LOG, "ImpostazioniIniziali: ATTENZIONE NESSUNA STAZIONE CONFIGURATA!")
                Me.Close()
            Else
                'Federica luglio 2018
                ReDim Linee(numeroLinee - 1)

                'Estraggo le linee presenti
                rsDati = DB.Reader("SELECT * FROM WDS_GENTAB WHERE GT_TYPE = 'STATIONS' ORDER BY GT_ORDER")
                For i = 0 To rsDati.Rows.Count - 1
                    Linee(i) = New Linea
                    Linee(i).CodiceLinea = Trim(CStr(rsDati.Rows(i)("GT_CODE")))
                    Linee(i).NumeroLinea = i + 1
                    Linee(i).NomeLinea = Trim(CStr(rsDati.Rows(i)("GT_DESCRIPTION")))
                Next

                'Imposto la prima come linea selezionata
                gsClienteDi = Linee(0).CodiceLinea
                gLineaSelzionata = 0

            End If

            rsDati = Nothing

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "GestioneLinee errore: " + ex.ToString)

        End Try

    End Sub

    'luca marzo 2018
    '    Sub GestioneQAL2QAL3()
    '        'Non viene chiamata perchè non abbiamo la QAL3
    '        Dim RetM As Object
    '        Dim RetQ As Object
    '        Dim RetIC As Object
    '        Dim RetRANGE As Object
    '        Dim RetData As Object

    '        On Error GoTo GestErrore

    '        'Call WindasLog(MSG_LOG, "START: " & Now)

    '        GestioneQAL2QAL3CaricaDatiQAL3(gsClienteDi, arrayMisureQAL2QAL3(0), txtZeroRif_0, txtSpanRif_0, txtDataUltimaQAL3_0, txtOraUltimaQAL3_0, txtZeroRis_0, txtSpanRis_0, txtRisultatoQAL3_0)
    '        GestioneQAL2QAL3CaricaDatiQAL3(gsClienteDi, arrayMisureQAL2QAL3(1), txtZeroRif_1, txtSpanRif_1, txtDataUltimaQAL3_1, txtOraUltimaQAL3_1, txtZeroRis_1, txtSpanRis_1, txtRisultatoQAL3_1)
    '        GestioneQAL2QAL3CaricaDatiQAL3(gsClienteDi, arrayMisureQAL2QAL3(2), txtZeroRif_2, txtSpanRif_2, txtDataUltimaQAL3_2, txtOraUltimaQAL3_2, txtZeroRis_2, txtSpanRis_2, txtRisultatoQAL3_2)
    '        'Call WindasLog(MSG_LOG, "STOP: " & Now)
    '        Exit Sub

    'GestErrore:
    '        Call WindasLog(ERR_LOG, "GestioneQAL2QAL3 errore: " + Err.Description)

    '    End Sub

    'luca marzo 2018
    Sub GestioneColorazioneSelettoriQAL3()

        On Error GoTo GestErrore

        GestioneSelettoriQAL(btnZero_1, "16")
        GestioneSelettoriQAL(btnZero_2, "17")
        GestioneSelettoriQAL(btnZero_3, "18")
        GestioneSelettoriQAL(btnZero_4, "19")

        GestioneSelettoriQAL(btnSpan_1, "24")
        GestioneSelettoriQAL(btnSpan_2, "25")
        GestioneSelettoriQAL(btnSpan_3, "26")
        GestioneSelettoriQAL(btnSpan_4, "27")

        Exit Sub

GestErrore:
        Call WindasLog(ERR_LOG, "GestioneColorazioneSelettoriQAL3 errore: " + Err.Description)

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            DataOra.Text = Format(Now, "dd MMMM yyyy HH.mm.ss")

            'Alby Maggio 2018 
            'Call GestioneSinottico()
            Call GestioneAllarmi()

            Call GestioneCasella("IST")
            Call GestioneCasella("ISTN")
            Call GestioneCasella("MONU")
            Call GestioneCasella("MONC")
            Call GestioneCasella("MONP")

            Call GestioneCasella("MGONU")
            Call GestioneCasella("MGONC")
            Call GestioneCasella("MGONP")

            'luca marzo 2018 gestione colorazione selettori QAL3 pagina QAL2/QAL3
            Call GestioneColorazioneSelettoriQAL3()

            'luca maggio 2018 stato impianto
            lblStatoImpianto1.Text = LeggiTag("1.STATO_IMPIANTO_STR")
            lblLinea1.Text = Linee(0).NomeLinea

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "Timer1 errore: " + ex.ToString)
        End Try

    End Sub

    'Sub GestioneSinottico()

    '    Dim Indice As String
    '    Dim Valore As Double
    '    Dim Oggetto
    '    Dim iIdx As Integer

    '    Dim ctrl As Control

    '    Try
    '        With Linee(gLineaSelzionata)
    '            For iIdx = 0 To .NrMisure
    '                Indice = .Misure(iIdx).Indice

    '                Inserimento dati in Pagina Sinottico
    '                ctrl = GetControlByName("ValSinottico" + Indice, TabSinottico)
    '                If Not ctrl Is Nothing Then
    '                    Valore = CDbl(LeggiTag(.NumeroLinea + ".AM" + Indice + "_IST"))
    '                    If Valore <> -9999 Then
    '                        ctrl.Text = Format(Valore, "0.00")
    '                    Else
    '                        ctrl.Text = "---"
    '                    End If
    '                End If
    '            Next iIdx

    '            For iIdx = 0 To .NrAllarmi
    '                Indice = Format(iIdx, "000")

    '                Allarmi()
    '                ctrl = GetControlByName("D" + Indice, TabAllarmi)
    '                If Not ctrl Is Nothing Then
    '                    Valore = Val(LeggiTag(.NumeroLinea + ".DM" + Indice + "_IST"))
    '                    ctrl.BackColor = ColoreSegnalazione(ctrl.Text, Valore)
    '                End If

    '                Lampade(sinottico)
    '                Oggetto = GetControlByName("Digital" + Indice, TabSinottico)
    '                If Not Oggetto Is Nothing Then
    '                    If Valore = 1 Then
    '                        Oggetto.image = WindasNet.My.Resources.Resources.KO
    '                        Oggetto.Visible = Not Oggetto.Visible
    '                    Else
    '                        Oggetto.image = WindasNet.My.Resources.Resources.OK
    '                        Oggetto.Visible = True
    '                    End If
    '                End If
    '            Next iIdx
    '        End With

    '    Catch ex As Exception
    '        Call WindasLog(ERR_LOG, "GestioneSinottico errore: " + ex.ToString)

    '    End Try

    'End Sub

    Sub GestioneAllarmi()

        Dim Indice As String
        Dim Valore As Double
        Dim iIdx As Integer
        Dim ctrl As Control

        Try
            With Linee(gLineaSelzionata)
                'Allarmi
                For iIdx = 0 To .NrAllarmi
                    'Indice = Format(.Allarmi(iIdx, 1), "000")
                    Indice = .Allarmi(iIdx, 1)

                    ' If Indice = "040" Then Stop

                    ctrl = GetControlByName("D" + Indice, TabAllarmi)
                    If Not ctrl Is Nothing Then
                        Valore = Val(LeggiTag(.NumeroLinea + ".DM" + Indice + "_IST"))
                        ctrl.BackColor = ColoreSegnalazione(Valore, 0)
                    End If
                Next iIdx

                'Stati
                For iIdx = 0 To .NrStati
                    'Indice = Format(iIdx, "000")
                    Indice = .Stati(iIdx, 1)

                    ctrl = GetControlByName("D" + Indice, TabStati)
                    If Not ctrl Is Nothing Then
                        Valore = Val(LeggiTag(.NumeroLinea + ".DM" + Indice + "_IST"))
                        ctrl.BackColor = ColoreSegnalazione(Valore, 1)
                    End If
                Next iIdx

                'Soglie
                For iIdx = 0 To .NrSoglie
                    'Indice = Format(iIdx, "000")
                    Indice = .Soglie(iIdx, 1)

                    ctrl = GetControlByName("D" + Indice, TabSoglie)
                    If Not ctrl Is Nothing Then
                        Valore = Val(LeggiTag(.NumeroLinea + ".DM" + Indice + "_IST"))
                        ctrl.BackColor = ColoreSegnalazione(Valore, 0)
                    End If
                Next iIdx

                'Calibrazioni
                For iIdx = 0 To .NrCalibrazioni
                    Indice = .Calibrazioni(iIdx)

                    ctrl = GetControlByName("lblQAL3_" + Indice, TabQAL)
                    If Not ctrl Is Nothing Then
                        Valore = Val(LeggiTag(.NumeroLinea + ".DM" + Indice + "_IST"))
                        ctrl.Visible = CBool(Valore)
                    End If
                Next iIdx
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "GestioneSinottico errore: " + ex.ToString)

        End Try

    End Sub

    Function ColoreSegnalazione(ByVal valore_segnalazione As Integer, ByVal tipo As Integer) As Color

        Try

            Select Case tipo
                Case 0  'Allarmi
                    If valore_segnalazione = 1 Then
                        ColoreSegnalazione = Color.Red

                    Else
                        ColoreSegnalazione = Color.DarkGreen
                    End If

                Case 1
                    If valore_segnalazione = 1 Then
                        ColoreSegnalazione = Color.DarkGreen

                    Else
                        ColoreSegnalazione = Color.Gray
                    End If
            End Select

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "ColoreSegnalazione: " & ex.ToString)
        End Try

    End Function

    Sub GestioneCasella(TipoTag)

        Dim Indice As String
        Dim Valore As Double
        Dim Validita As String
        Dim Oggetto
        Dim Supero As Integer
        Dim ctrl As Control
        Dim Percentuale As Integer

        Try
            With Linee(gLineaSelzionata)
                For iIdx = 0 To .NrMisure
                    Indice = .Misure(iIdx).Indice
                    ctrl = GetControlByName(TipoTag + Indice, TabMisure)
                    If Not ctrl Is Nothing Then
                        Valore = LeggiTag(.NumeroLinea + ".AM" + Indice + "_" + TipoTag)
                        Validita = LeggiTag(.NumeroLinea + ".AM" + Indice + "_" + TipoTag + "_VAL")

                        'luca maggio 2018
                        Supero = CInt(LeggiTag(.NumeroLinea + ".AM" + Indice + "_" + TipoTag + "_VIS"))

                        If Valore <> -9999 Then
                            ctrl.Text = Format(Valore, "0.00")
                        Else
                            ctrl.Text = "---"
                        End If

                        If Not IsNothing(Validita) Then
                            If Validita = "VAL" Then
                                ctrl.BackColor = System.Drawing.Color.White
                                If Supero = 5 Then
                                    ctrl.BackColor = System.Drawing.Color.Yellow
                                ElseIf Supero = 2 Then
                                    ctrl.BackColor = System.Drawing.Color.Red
                                End If
                            Else
                                ctrl.BackColor = System.Drawing.Color.DarkOrange
                            End If
                        Else
                            'Federica luglio 2018 - Le medie previsionali non hanno la validità
                            ctrl.BackColor = System.Drawing.Color.White
                            If Supero = 5 Then
                                ctrl.BackColor = System.Drawing.Color.Yellow
                            ElseIf Supero = 2 Then
                                ctrl.BackColor = System.Drawing.Color.Red
                            End If
                        End If
                    End If

                    If TipoTag = "IST" Then
                        Oggetto = GetControlByName("Gauge" + Indice, TabMisure)
                        If Not Oggetto Is Nothing Then
                            Valore = LeggiTag(.NumeroLinea + ".AM" + Indice + "_" + TipoTag)

                            If Valore <> -9999 Then
                                'luca maggio 2018
                                If Valore > .Misure(iIdx).FondoScala Then
                                    Percentuale = 100
                                ElseIf Valore < 0 Then
                                    Percentuale = 0
                                Else
                                    Percentuale = Valore / .Misure(iIdx).FondoScala * 100
                                    If Percentuale < 0 Then Percentuale = 0
                                End If
                            Else
                                Percentuale = 0
                            End If
                            Oggetto.Value = Percentuale
                        End If
                    End If
                Next
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "GestioneCasella errore: " + ex.ToString)

        End Try

    End Sub

    Dim WebBrowser_new As New ChromiumWebBrowser("http://localhost/pagine/sinotticowebform?wb=1")
    Public Sub New()

        ' Chiamata richiesta dalla finestra di progettazione.
        InitializeComponent()
        Try
            Dim asd As New CefSharp.CefSettings()
            asd.IgnoreCertificateErrors = True
            asd.WindowlessRenderingEnabled = True
            If CefSharp.Cef.IsInitialized = False Then
                CefSharp.Cef.Initialize(asd)
            End If
            WebBrowser_new = New ChromiumWebBrowser("http://localhost/pagine/sinotticowebform?wb=1")
            WebBrowser_new.Dock = System.Windows.Forms.DockStyle.Fill
            WebBrowser_new.Location = New System.Drawing.Point(3, 3)
            WebBrowser_new.MinimumSize = New System.Drawing.Size(20, 20)
            WebBrowser_new.Name = "WebBrowser1"
            WebBrowser_new.Size = New System.Drawing.Size(1866, 803)
            WebBrowser_new.TabIndex = 0
            TabSinottico.Controls.Add(WebBrowser_new)
            AddHandler WebBrowser_new.ConsoleMessage, AddressOf ConsoleMessage_Event


        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().

    End Sub

    Private Sub Label_click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '??? A cosa serve
        Dim L As Label = sender

        If L.BackColor = Color.Red Then
            L.BackColor = Color.DarkGreen
        Else
            L.BackColor = Color.Red
        End If

    End Sub

    Private Sub ConsoleMessage_Event(sender As Object, e As ConsoleMessageEventArgs)
        MessageBox.Show(e.Message + " " + e.Line.ToString())
    End Sub

    Private Sub txtM_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtM000.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtQ_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQ000.KeyPress
        InserisciSoloValoriNumerici(e)
    End Sub

    Private Sub txtQ_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQ011.KeyPress
        InserisciSoloValoriNumerici(e)
    End Sub

    Private Sub txtQ_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQ002.KeyPress
        InserisciSoloValoriNumerici(e)
    End Sub

    Private Sub txtQ_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e)
    End Sub

    Private Sub txtQ_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQ014.KeyPress
        InserisciSoloValoriNumerici(e)
    End Sub

    Private Sub txtM_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtM011.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtM_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtM002.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtM_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtM_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtM014.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtIC_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIC000.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtIC_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIC011.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtIC_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIC002.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtIC_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtIC_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIC014.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtRange_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRange000.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtRange_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRange011.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtRange_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRange002.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtRange_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtRange_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRange014.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtZeroRif_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtZeroRif_0.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtZeroRif_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtZeroRif_11.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtZeroRif_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtZeroRif_2.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtZeroRif_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtZeroRif_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSpanRis_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSpanRis_0.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSpanRis_1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSpanRis_11.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSpanRis_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSpanRis_2.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSpanRis_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSpanRis_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub TabQAL_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabQAL.Enter

        Call VisualizzaDati_TabQAL2QAL3()

    End Sub

    Private Sub UpdateQAL2QAL3(ByVal IndiceMisura As Integer)

        Dim ctrl As Control
        Dim ctrlData As DateTimePicker

        Try
            With Linee(gLineaSelzionata)
                'Leggo i valori dai campi
                ctrl = GetControlByName("txtM" & .Misure(IndiceMisura).Indice, TabQAL)
                If Not ctrl Is Nothing Then .tmpQAL2(0).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                ctrl = GetControlByName("txtQ" & .Misure(IndiceMisura).Indice, TabQAL)
                If Not ctrl Is Nothing Then .tmpQAL2(1).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                ctrl = GetControlByName("txtRange" & .Misure(IndiceMisura).Indice, TabQAL)
                If Not ctrl Is Nothing Then .tmpQAL2(2).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                ctrl = GetControlByName("txtIC" & .Misure(IndiceMisura).Indice, TabQAL)
                If Not ctrl Is Nothing Then .tmpQAL2(3).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                ctrlData = GetControlByName("dtpDataQAL2" & .Misure(IndiceMisura).Indice, TabQAL)
                If Not ctrlData Is Nothing Then .tmpQAL2(4).Valore = Format(ctrlData.Value, "yyyyMMdd")

                'Lancio il salvataggio
                .QAL2QAL3AggiornaDati(IndiceMisura)
            End With

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "UpdateQAL2QAL3 errore: " & ex.ToString)

        End Try

    End Sub

    Private Sub btnUpdateQAL2QAL3_0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateQAL2QAL3_0.Click

        If SonoTuttiNumerici(grpQAL2_0, "txt") Then Call UpdateQAL2QAL3(0)

    End Sub

    Private Sub btnUpdateQAL2QAL3_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateQAL2QAL3_1.Click

        If SonoTuttiNumerici(grpQAL2_1, "txt") Then Call UpdateQAL2QAL3(1)

    End Sub

    Private Sub btnUpdateQAL2QAL3_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateQAL2QAL3_2.Click

        If SonoTuttiNumerici(grpQAL2_2, "txt") Then Call UpdateQAL2QAL3(2)

    End Sub

    Private Sub btnUpdateQAL2QAL3_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateQAL2QAL3_4.Click

        If SonoTuttiNumerici(grpQAL2_4, "txt") Then Call UpdateQAL2QAL3(4)

    End Sub

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReport.Click

        Try
            Shell("c:\windas\BFReportClient\BFReportClient.exe", AppWinStyle.NormalFocus)

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "btnReport_Click errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub btnTrend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrend.Click

        Try
            Shell("C:\Windas\BFtrend\BFTrend.exe", AppWinStyle.NormalFocus)

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "btnTrend_Click errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub btnStartQAL3_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartQAL3_1.Click

        Call ScriviTagDOQAL("6")

    End Sub

    Private Sub btnStopQAL3_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopQAL3_1.Click

        Call ScriviTagDOQAL("7")

    End Sub

    Private Sub btnAutocalibrazione_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAutocalibrazione.Click

        Call ScriviTagDOQAL("12")

    End Sub

    Private Sub cmbLinea_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'luca marzo 2018
        Try
            gsClienteDi = Linee(CType(sender, ComboBox).SelectedIndex).CodiceLinea
            gLineaSelzionata = CType(sender, ComboBox).SelectedIndex

            'Federica luglio 2018
            'Non so in che pagina sono, quindi ricarico la visualizzazione di tutte
            Call VisualizzaDati_TabSinottico()
            Call VisualizzaDati_TabMisure()
            Call VisualizzaDati_TabAllarmi()
            Call VisualizzaDati_TabQAL2QAL3()
            Call VisualizzaDati_TabConfigurazione()

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "cmbLinea_SelectedIndexChanged errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub TabConfigurazione_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabConfigurazione.Enter

        Call VisualizzaDati_TabConfigurazione()

    End Sub

    Private Sub btnUpdateConf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateConf.Click

        Dim i As Integer
        Dim ctrl As Control
        Dim Indice As Integer

        Try
            'Controllo che tutti i valori inseriti siano numerici
            If SonoTuttiNumerici(TabConfigurazione, "txtSoglia") Then
                With Linee(gLineaSelzionata)
                    For i = 0 To .NrMisure
                        'Azzero le soglie temporanee
                        For Indice = 0 To UBound(.tmpSoglie)
                            .tmpSoglie(Indice).Valore = "0"
                        Next

                        'Leggo i valori dai campi
                        ctrl = GetControlByName("txtSogliaAtt_" & CStr(i), TabConfigurazione)
                        If Not ctrl Is Nothing Then .tmpSoglie(0).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                        ctrl = GetControlByName("txtSogliaAll_" & CStr(i), TabConfigurazione)
                        If Not ctrl Is Nothing Then .tmpSoglie(1).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                        ctrl = GetControlByName("txtSogliaAttGrn_" & CStr(i), TabConfigurazione)
                        If Not ctrl Is Nothing Then .tmpSoglie(2).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                        ctrl = GetControlByName("txtSogliaAllGrn_" & CStr(i), TabConfigurazione)
                        If Not ctrl Is Nothing Then .tmpSoglie(3).Valore = Val(Replace(Trim(ctrl.Text), ",", "."))

                        'Lancio il salvataggio
                        .ConfigurazioneAggiornaDati(i)
                    Next

                    'Alla fine rileggo la configurazione e aggiorno la visualizzazione
                    .CaricaMisure()
                    Call VisualizzaDati_TabConfigurazione()
                End With
            Else
                MsgBox("ATTENZIONE! I VALORI INSERITI DEVONO ESSERE NUMERICI! VERIFICARE!!!")
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "btnUpdateConf_Click errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub txtSogliaAtt_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_0.KeyPress, txtSogliaAtt_0.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAtt_1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAtt_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_3.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAtt_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_11.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_9.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_13.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_2.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_8.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAtt_4.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAtt_9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAll_0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_0.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAll_1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAll_2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_3.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub txtSogliaAll_3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_11.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_9.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_13.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_2.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_8.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSogliaAll_4.KeyPress
        InserisciSoloValoriNumerici(e, True)
    End Sub
    Private Sub txtSogliaAll_9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        InserisciSoloValoriNumerici(e, True)
    End Sub

    Private Sub TabMisure_Click(sender As System.Object, e As System.EventArgs) Handles TabMisure.Click

        Call VisualizzaDati_TabMisure()

    End Sub

    Private Sub TmrConfigurazione_Tick(sender As System.Object, e As System.EventArgs) Handles TmrConfigurazione.Tick

        'Rileggo la configurazione una volta al minuto
        Dim iIdx As Integer

        Try
            'Carico i dati per ciascuna linea
            For iIdx = 0 To UBound(Linee)
                Linee(iIdx).CaricaMisure()
                Linee(iIdx).CaricaAllarmi()
            Next

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "TmrConfigurazione_Tick errore: " + ex.ToString)

        End Try

    End Sub

    Private Sub btnUpdateQAL2QAL3_3_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdateQAL2QAL3_3.Click

        If SonoTuttiNumerici(grpQAL2_3, "txt") Then Call UpdateQAL2QAL3(3)

    End Sub

    Private Sub GestioneSelettoriQAL(ByRef btn As Button, ByVal sDO As String)

        Dim NomeTag As String
        Dim ValoreTag As String

        Try
            NomeTag = Linee(gLineaSelzionata).NumeroLinea.ToString & " DO" & sDO
            ValoreTag = LeggiTag(NomeTag)
            btn.BackColor = IIf(ValoreTag = 1, Color.ForestGreen, Color.Red)

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "GestioneSelettoriQAL: " & ex.ToString)

        End Try

    End Sub

    Private Sub ScriviTagDOQAL(ByVal sDO As String)

        Dim NomeTag As String

        Try
            NomeTag = Linee(gLineaSelzionata).NumeroLinea.ToString & " DO" & sDO
            Call ScriviTag(NomeTag, "1")

            'luca maggio 2018 - Impulsivo
            Threading.Thread.Sleep(1000)

            Call ScriviTag(NomeTag, "0")

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "ScriviTagDOQAL: " & ex.ToString)

        End Try

    End Sub

    Private Sub btnStartQAL3_2_Click_1(sender As System.Object, e As System.EventArgs) Handles btnStartQAL3_2.Click

        Call ScriviTagDOQAL("10")

    End Sub

    Private Sub btnStopQAL3_2_Click_1(sender As System.Object, e As System.EventArgs) Handles btnStopQAL3_2.Click

        Call ScriviTagDOQAL("11")

    End Sub

    Private Sub btnFermo_Click(sender As System.Object, e As System.EventArgs) Handles btnFermo.Click

        Call CambiaStatoImpianto(StatiImpianto.IMPIANTO_FERMO)

    End Sub

    Private Sub btnInMarcia_Click(sender As System.Object, e As System.EventArgs) Handles btnInMarcia.Click

        Call CambiaStatoImpianto(StatiImpianto.IMPIANTO_SOTTO_MINIMO_TECNICO)

    End Sub

    Private Sub btnSopra_Click(sender As System.Object, e As System.EventArgs) Handles btnSopra.Click

        Call CambiaStatoImpianto(StatiImpianto.IMPIANTO_IN_MARCIA)

    End Sub

    Private Sub btnSotto_Click(sender As System.Object, e As System.EventArgs) Handles btnSotto.Click

        Call CambiaStatoImpianto(StatiImpianto.IMPIANTO_SOTTO_MINIMO_TECNICO)

    End Sub

    Private Sub CambiaStatoImpianto(ByVal NuovoStato As StatiImpianto)

        Try

            Select Case NuovoStato
                Case StatiImpianto.IMPIANTO_FERMO
                    btnFermo.BackColor = Color.Red
                    btnInMarcia.BackColor = Color.Gray
                    btnSotto.BackColor = Color.Gray
                    btnSopra.BackColor = Color.Gray

                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI144", "0")
                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI145", "0")

                Case StatiImpianto.IMPIANTO_SOTTO_MINIMO_TECNICO

                    btnFermo.BackColor = Color.Gray
                    btnInMarcia.BackColor = Color.ForestGreen
                    btnSotto.BackColor = Color.Red
                    btnSopra.BackColor = Color.Gray

                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI144", "1")
                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI145", "1")

                Case StatiImpianto.IMPIANTO_IN_MARCIA
                    btnFermo.BackColor = Color.Gray
                    btnInMarcia.BackColor = Color.ForestGreen
                    btnSotto.BackColor = Color.Gray
                    btnSopra.BackColor = Color.ForestGreen

                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI144", "1")
                    Call ScriviTag(Linee(gLineaSelzionata).NumeroLinea.ToString & " DI145", "0")
            End Select

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "CambiaStatoImpianto: " & ex.ToString)
        End Try

    End Sub

    Private Enum StatiImpianto
        IMPIANTO_FERMO = 34
        IMPIANTO_IN_MARCIA = 30
        IMPIANTO_SOTTO_MINIMO_TECNICO = 31
    End Enum

End Class
