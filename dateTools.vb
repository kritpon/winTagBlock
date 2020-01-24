Public Class dateTools
    Function strToDate(strDate As String) As String
        Dim strDD As String
        Dim strMM As String
        Dim strYY As String
        Dim strDate2 As String
        Dim strTime As String


        ' trh_Date = Format(Month((.Range("H" & countRow + 1).Value)), "00") 
        '& "/" & Format(Microsoft.VisualBasic.Day((.Range("H" & countRow + 1).Value)), "00")
        '& "/" & (Year((.Range("H" & countRow + 1).Value)) - 543)

        'If Len(strDate) = 17 Then

        strMM = Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2)) 'Month(strDate) '
        strDD = Trim(Microsoft.VisualBasic.Left(strDate, 2)) 'Microsoft.VisualBasic.DateAndTime.Day(strDate) '
        strYY = Trim(Year(strDate))
        If CInt(strYY) > 2562 Then
            strYY = Str(Int(Year(Now)) - 543)
        Else
            strYY = Str(Int(Year(Now)) - 543)
        End If
        strTime = Microsoft.VisualBasic.Right(strDate, 8)
        strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        'Else
        '    strMM = (Microsoft.VisualBasic.Left(strDate, 2))
        '    strDD = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 5), 2)
        '    strYY = Year(strDate)
        '    If CInt(strYY) > 2562 Then
        '        strYY = Str(Int(Year(Now)) - 543)
        '    Else
        '        strYY = Str(Int(Year(Now)) - 543)
        '    End If
        '    strTime = Microsoft.VisualBasic.Right(strDate, 8)
        '    strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        'End If

        Return strDate2


    End Function
End Class
