Module Utili

    Public Sub GetConnectionParams()
        '***** Lettura parametri di configurazione database su file bfdesk.xml *****

        Dim FileXML As String = Application.StartupPath & "\connections.xml"

        Try
            GestDB = New GestioneDatabase.Gestione(FileXML)
            DB = GestDB.Database
        Catch ex As Exception
            Call WindasLog(ERR_LOG, "GetConnectionParam " + ex.ToString)
        End Try

    End Sub

    'luca marzo 2018
    Public Sub InserisciSoloValoriNumerici(e As System.Windows.Forms.KeyPressEventArgs, Optional ByVal AnchePositivi As Boolean = False)

        Dim NonAmmessi() As Integer

        Try
            NonAmmessi = IIf(AnchePositivi, {8, 44, 46}, {8, 44, 46, 45})
            If Not NonAmmessi.Contains(Asc(e.KeyChar)) Then
                If Not Char.IsDigit(e.KeyChar) Then
                    e.Handled = True
                End If
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "InserisciSoloValoriNumerici errore: " + ex.ToString)

        End Try

    End Sub

    'Federica luglio 2018
    Public Function GetControlByName(ByVal Nome As String, ByVal Padre As Control) As Control

        Dim Ret() As Control

        Try
            GetControlByName = Nothing
            Ret = Padre.Controls.Find(Nome, True)
            If Ret.Count > 0 Then
                GetControlByName = Ret(0)
            End If

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "ParametriIniziali errore: " + ex.ToString)
            GetControlByName = Nothing

        End Try

    End Function

    'Federica luglio 2018
    Public Function SonoTuttiNumerici(ByRef Padre As Control, ByVal NomeControllo As String) As Boolean

        Dim oggetto As Control

        Try
            'Controllo che tutti i valori inseriti siano numerici
            SonoTuttiNumerici = True
            For Each ctl As Control In Padre.Controls
                If TypeOf ctl Is TextBox Then
                    oggetto = CType(ctl, TextBox)
                    If oggetto.Name.Substring(0, Len(NomeControllo)) = NomeControllo Then
                        If Not IsNumeric(ctl.Text) Then
                            SonoTuttiNumerici = False
                            Exit For
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            Call WindasLog(ERR_LOG, "SonoTuttiNumerici errore: " & ex.ToString)
            SonoTuttiNumerici = False

        End Try

    End Function

End Module
