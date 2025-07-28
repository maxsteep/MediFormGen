Attribute VB_Name = "invoiceNumGen"
Sub generateInvoiceNumbers()

    givenDate = Range("E2").Value
    givenInvoiceNumber = Range("F2").Value
    givenOrderNumber = Range("G2").Value
    workdays = Range("H2").Value
    monthBack = Range("I2").Value
    
    givenMonth = month(startingDate)
    givenYear = year(startingDate)
    arrayofWorkdays = Split(workdays, ",")
    
    startingDate = DateAdd("m", monthBack * -1, givenDate)
    endingDate = DateAdd("yyyy", 2, givenDate)
    
    currentDate = startingDate
    currentRow = 1
    flag = True
    
    Do While currentDate <> endingDate
        If isAWorkDay(CDate(currentDate), CStr(workdays)) Then
            Cells(currentRow, 1) = currentDate
            Cells(currentRow, 2) = 0
            If flag Then
                If currentDate >= givenDate Then
                    Cells(currentRow, 3).Value = givenInvoiceNumber
                    If givenOrderNumber <> "" Then
                        Cells(currentRow, 4).Value = givenOrderNumber
                    End If
                    flag = False
                Else
                    Cells(currentRow, 3).Formula = "=C" & currentRow + 1 & " - 25"
                    If givenOrderNumber <> "" Then
                        Cells(currentRow, 4).Formula = "=D" & currentRow + 1 & " - 25"
                    End If
                End If
            Else
                Cells(currentRow, 3).Formula = "=C" & currentRow - 1 & " + 25"
                If givenOrderNumber <> "" Then
                    Cells(currentRow, 4).Formula = "=D" & currentRow - 1 & " + 25"
                End If
            End If
            currentRow = currentRow + 1
        End If
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    
End Sub


Function isAWorkDay(givenDate As Date, workdays As String)

    Dim holidaysArray(1 To 63) As String
    'month/day/year
    holidaysArray(1) = "1/1/2014"
    holidaysArray(2) = "2/17/2014"
    holidaysArray(3) = "4/18/2014"
    holidaysArray(4) = "5/19/2014"
    holidaysArray(5) = "7/1/2014"
    holidaysArray(6) = "9/1/2014"
    holidaysArray(7) = "10/13/2014"
    holidaysArray(8) = "12/25/2014"
    holidaysArray(9) = "12/26/2014"
    
    holidaysArray(10) = "1/1/2015"
    holidaysArray(11) = "2/16/2015"
    holidaysArray(12) = "4/3/2015"
    holidaysArray(13) = "5/18/2015"
    holidaysArray(14) = "7/1/2015"
    holidaysArray(15) = "9/7/2015"
    holidaysArray(16) = "10/12/2015"
    holidaysArray(17) = "12/25/2015"
    holidaysArray(18) = "12/26/2015"
    
    holidaysArray(19) = "1/1/2016"
    holidaysArray(20) = "2/15/2016"
    holidaysArray(21) = "3/25/2016"
    holidaysArray(22) = "5/23/2016"
    holidaysArray(23) = "7/1/2016"
    holidaysArray(24) = "9/5/2016"
    holidaysArray(25) = "10/10/2016"
    holidaysArray(26) = "12/25/2016"
    holidaysArray(27) = "12/26/2016"
    
    holidaysArray(28) = "1/1/2017"
    holidaysArray(29) = "2/20/2017"
    holidaysArray(30) = "4/14/2017"
    holidaysArray(31) = "5/22/2017"
    holidaysArray(32) = "7/1/2017"
    holidaysArray(33) = "9/4/2017"
    holidaysArray(34) = "10/9/2017"
    holidaysArray(35) = "12/25/2017"
    holidaysArray(36) = "12/26/2017"
    
    holidaysArray(37) = "1/1/2018"
    holidaysArray(38) = "2/19/2018"
    holidaysArray(39) = "3/30/2018"
    holidaysArray(40) = "5/21/2018"
    holidaysArray(41) = "7/2/2018"
    holidaysArray(42) = "9/3/2018"
    holidaysArray(43) = "10/8/2018"
    holidaysArray(44) = "12/25/2018"
    holidaysArray(45) = "12/26/2018"
    
    holidaysArray(46) = "1/1/2019"
    holidaysArray(47) = "2/18/2019"
    holidaysArray(48) = "4/19/2019"
    holidaysArray(49) = "5/20/2019"
    holidaysArray(50) = "7/1/2019"
    holidaysArray(51) = "9/2/2019"
    holidaysArray(52) = "10/14/2019"
    holidaysArray(53) = "12/25/2019"
    holidaysArray(54) = "12/26/2019"

    holidaysArray(55) = "1/1/2020"
    holidaysArray(56) = "2/17/2020"
    holidaysArray(57) = "4/10/2020"
    holidaysArray(58) = "5/18/2020"
    holidaysArray(59) = "7/1/2020"
    holidaysArray(60) = "9/7/2020"
    holidaysArray(61) = "10/12/2020"
    holidaysArray(62) = "12/25/2020"
    holidaysArray(63) = "12/26/2020"
    
    arrayofWorkdays = Split(workdays, ",")
    isAWorkDay = (UBound(Filter(arrayofWorkdays, CStr(Weekday(givenDate)))) > -1) And Not (UBound(Filter(holidaysArray, CStr(givenDate))) > -1)
        
End Function
