Module VariabiliPubbliche

    'Federica luglio 2018
    Public Structure Misura
        Dim Indice As String
        Dim Codice As String
        Dim FondoScala As Double
        Dim Descrizione As String
        Dim ZeroRif As String
        Dim SpanRif As String
        Dim Soglie() As Campo
        Dim QAL2() As Campo

        Public Function GetSogliaByChiave(ByVal key As String) As Campo

            GetSogliaByChiave = Nothing
            For Each s In Soglie
                If s.Chiave = key Then
                    GetSogliaByChiave = s
                    Exit For
                End If
            Next
        End Function

        Public Function GetSogliaByCampoDB(ByVal key As String) As Campo

            GetSogliaByCampoDB = Nothing
            For Each s In Soglie
                If s.Campo = key Then
                    GetSogliaByCampoDB = s
                    Exit For
                End If
            Next
        End Function

        Public Function GetQAL2ByChiave(ByVal key As String) As Campo

            GetQAL2ByChiave = Nothing
            For Each s In QAL2
                If s.Chiave = key Then
                    GetQAL2ByChiave = s
                    Exit For
                End If
            Next

        End Function

        Public Function GetQAL2ByCampoDB(ByVal key As String) As Campo

            GetQAL2ByCampoDB = Nothing
            For Each s In QAL2
                If s.Campo = key Then
                    GetQAL2ByCampoDB = s
                    Exit For
                End If
            Next
        End Function

    End Structure

    Public Linee() As Linea

    '*** DATI PER CONNESSIONE A DATABASE ***
    'Federica gennaio 2018 - Nuova gestione connessioni
    Public Structure ConnectionsDB
        Dim StationCode As String
        Dim AppServer As String
        Dim AppDatabase As String
        Dim AppDBType As String
        Dim AppDBUser As String
        Dim AppDBPwd As String
        Dim AppScheduleWorking As Boolean
        Dim AppDbVersion As String
        Dim AppRS As Object
        Dim AppOrderSAD As Integer
        Dim AppDefaultDB As Boolean
    End Structure

    Public ConnessioneValida As Boolean
    Public connDB() As ConnectionsDB
    Public iConnDBDefault As Integer

    Public gsClienteDi As String

    Public Structure Campo
        Dim Campo As String
        Dim Valore
        Dim Chiave As String
    End Structure

    Public GestDB As GestioneDatabase.Gestione
    Public DB As GestioneDatabase.Database

End Module
