'Alby Febbraio 2018
Imports System
Imports System.IO

Module WinDAS

    Public Const MSG_LOG = 0
    Public Const ERR_LOG = 1
    Dim Comunicator As Object

    Public Sub InizializzaComunicator()

        Try
            'Alby Febbraio 2018
            Comunicator = GetObject("", "BFcomunicator.cloggerdata")
            Comunicator.startdatasharing()
        Catch ex As Exception
            Call WindasLog(ERR_LOG, "InizializzaComunicator errore: " + ex.ToString)
        End Try

    End Sub

    Function LeggiTag(TagName As String)

        Try
            Comunicator.CurrentItem = TagName
            LeggiTag = Comunicator.ItemValue
        Catch ex As Exception
            Call WindasLog(ERR_LOG, "LeggiTag errore: " + ex.ToString)
            LeggiTag = Nothing
        End Try

    End Function

    Sub ScriviTag(TagName As String, Valore As String)

        Try
            Comunicator.AddItem(TagName)
            Comunicator.CurrentItem = TagName
            Comunicator.ItemValue = Valore
        Catch ex As Exception
            Call WindasLog(ERR_LOG, "ScriviTag errore: " + ex.ToString)
        End Try

    End Sub

    Public Sub WindasLog(ByVal tipo_log As Long, ByVal msg As String)

        Dim log_path As String
        Dim log_file As String = ""

        Try
            '***** crea cartella se non esiste ******
            log_path = Application.StartupPath & "\Log"
            If (Dir(log_path, vbDirectory) = "") Then MkDir(log_path)

            '********* seleziona file log ***********
            Select Case tipo_log
                Case ERR_LOG
                    log_file = log_path & "\" & Format(Now, "yyyyMMdd") & "_error.log"

                Case MSG_LOG
                    log_file = log_path & "\" & Format(Now, "yyyyMMdd") & "_activity.log"
            End Select

            Using sw As StreamWriter = File.AppendText(log_file)
                sw.WriteLine(Format(Now, "dd/MM/yyyy HH.mm.ss") & Chr(9) & msg)
            End Using
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

End Module



