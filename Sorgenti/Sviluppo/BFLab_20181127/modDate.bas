Attribute VB_Name = "modDate"
Public Function DateTimeSerial(year As Integer, month As Integer, day As Integer, hour As Integer, minute As Integer, second As Integer) As Date
    DateTimeSerial = DateSerial(year, month, day) + TimeSerial(hour, minute, second)
End Function

'luca 06.03.2015 funzione per modificare l'anno di una data (senza trasformare la data in stringa)
Public Function SetYear(Data As Date, year As Integer) As Date
    SetYear = DateTimeSerial(year, month(Data), day(Data), hour(Data), minute(Data), second(Data))
End Function

'luca 06.03.2015 funzione per modificare il mese di una data (senza trasformare la data in stringa)
Public Function SetMonth(Data As Date, month As Integer) As Date
    If month >= 1 And month <= 12 Then
        SetMonth = DateTimeSerial(year(Data), month, day(Data), hour(Data), minute(Data), second(Data))
    Else
        Call Err.Raise(vbObjectError + 1, "BFlabServer.SetMonth", "Mese non valido: " & CStr(month))
    End If
End Function

'luca 06.03.2015 funzione per modificare il giorno di una data (senza trasformare la data in stringa)
'                implementata gestione modifica del giorno in base al mese (introdotta gestione anno bisestile)
Public Function SetDay(Data As Date, day As Integer) As Date
    
    Select Case month(Data)
    
        Case 1, 3, 5, 7, 8, 10, 12
            If day >= 1 And day <= 31 Then
                SetDay = DateTimeSerial(year(Data), month(Data), day, hour(Data), minute(Data), second(Data))
            Else
                GoTo Errore
            End If
            
        Case 2
            'luca 06.03.2015 gestione anni bisestili
            'Un anno è bisestile se il suo numero è divisibile per 4, con l'eccezione degli anni secolari (quelli divisibili per 100) che non sono divisibili per 400.
            If (year(Data) Mod 400 = 0 Or (year(Data) Mod 100 <> 0 And year(Data) Mod 4 = 0)) Then
                If day >= 1 And day <= 29 Then
                    SetDay = DateTimeSerial(year(Data), month(Data), day, hour(Data), minute(Data), second(Data))
                Else
                    GoTo Errore
                End If
            Else
                If day >= 1 And day <= 28 Then
                    SetDay = DateTimeSerial(year(Data), month(Data), day, hour(Data), minute(Data), second(Data))
                Else
                    GoTo Errore
                End If
            End If
            
        Case 4, 6, 9, 11
            If day >= 1 And day <= 30 Then
                SetDay = DateTimeSerial(year(Data), month(Data), day, hour(Data), minute(Data), second(Data))
            Else
                GoTo Errore
            End If
            
    End Select
    
Exit Function

Errore:
    Call Err.Raise(vbObjectError + 1, "BFlabServer.SetDay", "Giorno " & CStr(day) & " del mese " & month(Data) & " non valido ")
    
End Function

'luca 06.03.2015 funzione per modificare l'ora di una data (senza trasformare la data in stringa)
Public Function SetHour(Data As Date, hour As Integer) As Date
    If hour >= 0 And hour <= 23 Then
        SetHour = DateTimeSerial(year(Data), month(Data), day(Data), hour, minute(Data), second(Data))
    Else
        Call Err.Raise(vbObjectError + 1, "BFlabServer.SetHour", "Ore non valide: " & CStr(hour))
    End If
End Function

'luca 06.03.2015 funzione per modificare il minuto di una data (senza trasformare la data in stringa)
Public Function SetMinute(Data As Date, minute As Integer) As Date
    If minute >= 0 And minute <= 59 Then
        SetMinute = DateTimeSerial(year(Data), month(Data), day(Data), hour(Data), minute, second(Data))
    Else
        Call Err.Raise(vbObjectError + 1, "BFlabServer.SetMinute", "Minuti non validi: " & CStr(minute))
    End If
End Function

'luca 06.03.2015 funzione per modificare il secondo di una data (senza trasformare la data in stringa)
Public Function SetSecond(Data As Date, second As Integer) As Date
    If second >= 0 And second <= 59 Then
        SetSecond = DateTimeSerial(year(Data), month(Data), day(Data), hour(Data), minute(Data), second)
    Else
        Call Err.Raise(vbObjectError + 1, "BFlabServer.SetSecond", "Secondi non validi: " & CStr(second))
    End If
End Function

'luca 09.03.2015 funzione per creare una data partendo dalla data estratta dal DB (quindi con formato yyyymmdd)
Public Function CreateDateFromDB(DateFromDB As String) As Date
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    
    year = CInt(Mid(DateFromDB, 1, 4))
    month = CInt(Mid(DateFromDB, 5, 2))
    day = CInt(Mid(DateFromDB, 7, 2))
    
    CreateDateFromDB = DateSerial(year, month, day)
    'luca 09.03.2015
    'errore non gestito perchè non è possibile determinare il motivo per cui la data non viene convertita correttamente
    'in caso di errore esso è da gestire a livello di chiamata della funzione
    
End Function

'luca 09.03.2015 funzione per creare un'ora partendo dall'ora estratta dal DB (quindi con formato hh.nn.ss)
Public Function CreateTimeFromDB(TimeFromDB As String) As Date
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer

    hour = CInt(Mid(TimeFromDB, 1, 2))
    minute = CInt(Mid(TimeFromDB, 4, 2))
    second = CInt(Mid(TimeFromDB, 7, 2))
    
    CreateTimeFromDB = TimeSerial(hour, minute, second)
    
    'luca 09.03.2015
    'errore non gestito perchè non è possibile determinare il motivo per cui la data non viene convertita correttamente
    'in caso di errore esso è da gestire a livello di chiamata della funzione
    
End Function

'luca 09.03.2015 funzione per creare una data (comprensiva di data e ora) partendo da data e ora estratte dal DB (quindi con formato yyyymmdd e hh.nn.ss)
Public Function CreateDateTimeFromDB(DateFromDB As String, TimeFromDB As String) As Date
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    
    Dim hour As Integer
    Dim minute As Integer
    Dim second As Integer
    
    year = CInt(Mid(DateFromDB, 1, 4))
    month = CInt(Mid(DateFromDB, 5, 2))
    day = CInt(Mid(DateFromDB, 7, 2))
    
    hour = CInt(Mid(TimeFromDB, 1, 2))
    minute = CInt(Mid(TimeFromDB, 4, 2))
    second = CInt(Mid(TimeFromDB, 7, 2))
    
    CreateDateTimeFromDB = DateSerial(year, month, day) + TimeSerial(hour, minute, second)
    
    'luca 09.03.2015
    'errore non gestito perchè non è possibile determinare il motivo per cui la data non viene convertita correttamente
    'in caso di errore esso è da gestire a livello di chiamata della funzione
    
End Function

Public Function CreateDateForDB(DateForDB As Date) As String
    CreateDateForDB = Format(DateForDB, "yyyymmdd")
End Function
Public Function CreateDateForSQLS(DateForDB As Date) As String
    CreateDateForSQLS = Format(DateForDB, "yyyy-MM-ddTHH:mm:ss")
End Function
