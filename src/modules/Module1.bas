Attribute VB_Name = "Module1"
' Add a reference to the Word-library via VBE > Tools > References > Microsoft Word xx.x Object Library.
' STRIPPED_EXAMPLE_PASSWORD is a PLACEHOLDER not a real, nor ever used password.
' logging and auditing was performed separately

'''Word File Variables
Public wrdApp As Word.Application
Public wrdDoc As Word.Document
Public bWeStartedWord As Boolean
'''Global Vars
Public formMonth As String
Public formName As String
Public formYear As String
Public formDate As String

Public errorMSG As String

Public wholeFilePath As String

Public selectedForms(6) As Boolean

'Function StripChar(s As String) As String
'    With CreateObject("vbscript.regexp")
'        .Global = True
'        .ignorecase = True
'        .Pattern = "[^\dA-Z ]"
'        StripChar = .Replace(s, "")
'    End With
'End Function


Sub openWordDocument(bluePrintPath As String)
    On Error Resume Next
    Set wrdApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If wrdApp Is Nothing Then
        Set wrdApp = CreateObject("Word.Application")
        bWeStartedWord = True
    End If
    wrdApp.Visible = True
    Set wrdDoc = wrdApp.Documents.Open(bluePrintPath, PasswordDocument:="STRIPPED_EXAMPLE_PASSWORD")
End Sub

Sub replaceText(objDoc As Word.Document, findString As String, replaceString As String)
    Dim rngStory As Object
     For Each rngStory In objDoc.StoryRanges
        With rngStory.Find
            .Text = findString
            .Replacement.Text = replaceString
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next

End Sub


Sub massGenerate()
    Dim wholePath As String
    Dim i As Integer
    Dim rawInvoiceNumber As String
    Dim NAME As String
    Dim FORMmainDate As String
    Dim Doc As String
    Dim ADDRESS1 As String
    Dim ADDRESS2 As String
    Dim NAME2 As String
    Dim DOB As String
    Dim IDNO As String
    Dim GROUPNO As String
    Dim orderNo As String
    Dim INVOICENO As String
    Dim formDateIndex, rawLength As Integer
    Dim specialDay1 As String
    Dim specialDay2 As String
    
    Dim monthYear As String
    
    Dim numberOfDays As Integer
    Dim workingDate As Date
    
    Dim docList
    Dim docCounter As Integer

    errorMSG = ""
    
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect ("STRIPPED_EXAMPLE_PASSWORD")

    For i = 0 To 5
        selectedForms(i) = ThisWorkbook.Worksheets("Control").Shapes("Check Box " & (i + 1)).OLEFormat.Object.Value = 1
    Next i
    i = 0
    If checkSelectedFormVals Then
        If areThereAvaliableDates Then
            For i = 0 To 5
                On Error GoTo cleanUpAllOnError
                If selectedForms(i) Then
                    docList = Split(getDoc(i), "|")
                    docCounter = 0
                    For Each Docx In docList
                        Doc = CStr(Docx)
                        docCounter = docCounter + 1
                        If isFilled(Doc) Then
                                                
                            wholePath = CStr(Application.ActiveWorkbook.Path & "\" & "Blue Prints\" & Doc & ".docx")
                            
                            Call openWordDocument(CStr(wholePath))
                            
                            ThisWorkbook.Sheets(CStr(Doc)).Visible = True
                            
                            NAME = Trim(ThisWorkbook.Sheets("Control").Range("C3").Value)
                            ADDRESS1 = Trim(ThisWorkbook.Sheets("Control").Range("C4").Value)
                            ADDRESS2 = Trim(ThisWorkbook.Sheets("Control").Range("C5").Value)
                            NAME2 = Trim(ThisWorkbook.Sheets("Control").Range("C7").Value)
                            DOB = Trim(ThisWorkbook.Sheets("Control").Range("C8").Value)
                            IDNO = Trim(ThisWorkbook.Sheets("Control").Range("C9").Value)
                            GROUPNO = Trim(ThisWorkbook.Sheets("Control").Range("C10").Value)
                            monthYear = Trim(ThisWorkbook.Sheets("Control").Range("C6").Value)
                            Select Case i
                                Case 0
                                    formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I2").Value), CStr(ThisWorkbook.Sheets("Control").Range("J2").Value))
                                Case 1
                                    formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I4").Value), CStr(ThisWorkbook.Sheets("Control").Range("J4").Value))
                                Case 2
                                    Select Case docCounter
                                        Case 1
                                            formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I7").Value), CStr(ThisWorkbook.Sheets("Control").Range("J7").Value))
                                        Case 2
                                            formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I8").Value), CStr(ThisWorkbook.Sheets("Control").Range("J8").Value))
                                    End Select
                                Case 3
                                    Select Case docCounter
                                        Case 1
                                            formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I10").Value), CStr(ThisWorkbook.Sheets("Control").Range("J10").Value))
                                        Case 2
                                            formDateIndex = getFormDateIndex(CStr(Doc), CStr(ThisWorkbook.Sheets("Control").Range("I11").Value), CStr(ThisWorkbook.Sheets("Control").Range("J11").Value))
                                    End Select
                                Case 4
                                    Select Case docCounter
                                        Case 1
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L14")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M14")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K14").Value))
                                        Case 2
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L15")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M15")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K15").Value))
                                        Case 3
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L16")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M16")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K16").Value))
                                    End Select
                                Case 5
                                    Select Case docCounter
                                        Case 1
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L19")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M19")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K19").Value))
                                        Case 2
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L20")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M20")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K20").Value))
                                        Case 3
                                            specialDay1 = ThisWorkbook.Sheets("Control").Range("L21")
                                            specialDay2 = ThisWorkbook.Sheets("Control").Range("M21")
                                            formDateIndex = getFormDateIndexSpecial(CStr(Doc), CDate(ThisWorkbook.Sheets("Control").Range("K21").Value))
                                    End Select
                            End Select
                            If formDateIndex <> -1 Then '' should never be -1 the check should catch
                                FORMmainDate = ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 1).Value
                                orderNo = CStr(CLng(ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 4).Value) + CLng(ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value))
                                rawInvoiceNumber = CStr(CLng(ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 3).Value) + CLng(ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value))
                                rawLength = CInt(ThisWorkbook.Sheets(Doc).Range("J2").Value)
                                INVOICENO = generateInvoiceNumber(i, rawInvoiceNumber, rawLength, NAME, FORMmainDate, Doc)
                            End If
                            
                            Select Case i
                                Case 0
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "mm/dd/yyyy"))
                                    Call replaceText(wrdDoc, "<<DATE+>>", Format(FORMmainDate, "mmmm dd, yyyy"))
                                    Call replaceText(wrdDoc, "<<NAME>>", getTitle(i) & NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                Case 1
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "mm/dd/yyyy"))
                                    Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                    Call replaceText(wrdDoc, "<<NAME2>>", NAME2)
                                    Call replaceText(wrdDoc, "<<DOB>>", Format(DOB, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<ID>>", IDNO)
                                    Call replaceText(wrdDoc, "<<GROUP>>", GROUPNO)
                                    Call replaceText(wrdDoc, "<<ORDER>>", orderNo)
                                Case 2
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<DATES1>>", Format(FORMmainDate, "dd-mmm-yyyy"))
                                    Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                Case 3
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<DATES1>>", Format(FORMmainDate, "dd-mmm-yyyy"))
                                    Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                    Call replaceText(wrdDoc, "<<NAME2>>", NAME2)
                                    Call replaceText(wrdDoc, "<<DOB>>", Format(DOB, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<ID>>", IDNO)
                                    Call replaceText(wrdDoc, "<<GROUP>>", GROUPNO)
                                Case 4
                                    Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                    numberOfDays = 15
                                    workingDate = CDate(FORMmainDate)
                                    Do While numberOfDays <> 0
                                        If Format(workingDate, "dddd") = specialDay1 Or Format(workingDate, "dddd") = specialDay2 And isAWorkDay(workingDate, "2,3,4,5,6") Then
                                            Call replaceText(wrdDoc, "<<DATES" & CStr(numberOfDays) & ">>", Format(workingDate, "dd-mmm-yyyy"))
                                            numberOfDays = numberOfDays - 1
                                        End If
                                        workingDate = DateAdd("d", -1, workingDate)
                                    Loop
                                Case 5
                                    Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                    Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                    Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                    Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                    Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<NAME2>>", NAME2)
                                    Call replaceText(wrdDoc, "<<DOB>>", Format(DOB, "dd/mmm/yyyy"))
                                    Call replaceText(wrdDoc, "<<ID>>", IDNO)
                                    Call replaceText(wrdDoc, "<<GROUP>>", GROUPNO)
                                    numberOfDays = 15
                                    workingDate = CDate(FORMmainDate)
                                    Do While numberOfDays <> 0
                                        If Format(workingDate, "dddd") = specialDay1 Or Format(workingDate, "dddd") = specialDay2 And isAWorkDay(workingDate, "2,3,4,5,6") Then
                                            Call replaceText(wrdDoc, "<<DATES" & CStr(numberOfDays) & ">>", Format(workingDate, "dd-mmm-yyyy"))
                                            numberOfDays = numberOfDays - 1
                                        End If
                                        workingDate = DateAdd("d", -1, workingDate)
                                    Loop
                            End Select
                            
                            Call saveAndClose(NAME, ADDRESS1, Doc, FORMmainDate, monthYear)
                            'On Error GoTo databaseError
                            On Error GoTo 0
                            Select Case i
                                Case 0 To 6
                                    With ThisWorkbook.Sheets("Database")
                                        .Range("A2").EntireRow.Insert
                                        .Rows(2).ClearFormats
                                        .Rows(2).RowHeight = 20
                                        
                                        .Cells(2, 1).Value = NAME
                                        .Cells(2, 2).Value = ADDRESS1
                                        .Cells(2, 3).Value = ADDRESS2
                                        
                                        .Cells(2, 4).Value = NAME2
                                        .Cells(2, 5).Value = DOB
                                        .Cells(2, 6).Value = IDNO
                                        .Cells(2, 7).Value = GROUPNO
                                                                        
                                        .Cells(2, 8).Value = FORMmainDate
                                        .Cells(2, 9).Hyperlinks.Add ThisWorkbook.Sheets("Database").Cells(2, 9), wholeFilePath, TextToDisplay:=INVOICENO
                                        
                                        .Cells(2, 10).Value = Trim(ThisWorkbook.Sheets("Control").Range("C2").Value)
                                        .Cells(2, 13).Value = monthYear
                                        '.Cells(2, 14).Value = i
                                        '.Cells(2, 15).Value = docCounter
                                    End With
                            End Select
                            Select Case i
                                Case 0
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(255, 255, 255)
                                Case 1
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(230, 230, 230)
                                Case 2
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(230, 230, 255)
                                Case 3
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(207, 207, 230)
                                Case 4
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(255, 230, 230)
                                Case 5
                                    ThisWorkbook.Sheets("Database").Range("A2:M2").Interior.Color = RGB(230, 207, 207)
                            End Select
                            Select Case i
                                Case 1
                                    ThisWorkbook.Sheets("Database").Cells(2, 11) = Trim(orderNo)
                                Case 4
                                    Select Case docCounter
                                        Case 1
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M14").Value)
                                        Case 2
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M15").Value)
                                        Case 3
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M16").Value)
                                    End Select
                                Case 5
                                    Select Case docCounter
                                        Case 1
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M19").Value)
                                        Case 2
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M20").Value)
                                        Case 3
                                            ThisWorkbook.Sheets("Database").Cells(2, 12) = Trim(ThisWorkbook.Sheets("Control").Range("M21").Value)
                                    End Select
                            End Select
                            ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value = ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value + 1
                            ThisWorkbook.Sheets(CStr(Doc)).Visible = 2
                            ThisWorkbook.Save
                            If Doc = "Chiropod - BRAMROSE" Then
                                On Error GoTo cleanUpAllOnError
                                Doc = "Orthotics - BRAMROSE"
                                wholePath = CStr(Application.ActiveWorkbook.Path & "\" & "Blue Prints\" & Doc & ".docx")
                                Call openWordDocument(CStr(wholePath))
                                ThisWorkbook.Sheets(CStr(Doc)).Visible = True
                                INVOICENO = generateInvoiceNumber(i, CStr(CLng(rawInvoiceNumber) + 1), rawLength, NAME, FORMmainDate, Doc)
                                Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                Call replaceText(wrdDoc, "<<DATES1>>", Format(FORMmainDate, "dd-mmm-yyyy"))
                                Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                Call replaceText(wrdDoc, "<<NAME2>>", NAME2)
                                Call replaceText(wrdDoc, "<<DOB>>", Format(DOB, "dd/mmm/yyyy"))
                                Call replaceText(wrdDoc, "<<ID>>", IDNO)
                                Call replaceText(wrdDoc, "<<GROUP>>", GROUPNO)
                                Call saveAndClose(NAME, ADDRESS1, Doc, FORMmainDate, monthYear)
                                On Error GoTo databaseError
                                '''ADD TO DATABASE
                                With ThisWorkbook.Sheets("Database")
                                    .Range("A2").EntireRow.Insert
                                    .Rows(2).ClearFormats
                                    .Rows(2).RowHeight = 20
                                    
                                    .Cells(2, 1).Value = NAME
                                    .Cells(2, 2).Value = ADDRESS1
                                    .Cells(2, 3).Value = ADDRESS2
                                    
                                    .Cells(2, 4).Value = NAME2
                                    .Cells(2, 5).Value = DOB
                                    .Cells(2, 6).Value = IDNO
                                    .Cells(2, 7).Value = GROUPNO
                                    
                                    .Cells(2, 8).Value = FORMmainDate
                                    .Cells(2, 9).Hyperlinks.Add .Cells(2, 9), wholeFilePath, TextToDisplay:=INVOICENO
                                    .Cells(2, 10).Value = Trim(ThisWorkbook.Sheets("Control").Range("C2").Value)
                                    .Cells(2, 13).Value = monthYear
                                    '.Cells(2, 14).Value = i
                                    '.Cells(2, 15).Value = docCounter
                                    .Range("A2:M2").Interior.Color = RGB(230, 230, 255)
                                End With
                                '''added to database
                                ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value = ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value + 1
                                ThisWorkbook.Sheets(CStr(Doc)).Visible = 2
                                ThisWorkbook.Save
                            End If
                            If Doc = "Chiropod - WALKIN COMFORT" Then
                                On Error GoTo cleanUpAllOnError
                                Doc = "Orthotics - WALKIN COMFORT"
                                wholePath = CStr(Application.ActiveWorkbook.Path & "\" & "Blue Prints\" & Doc & ".docx")
                                Call openWordDocument(CStr(wholePath))
                                ThisWorkbook.Sheets(CStr(Doc)).Visible = True
                                INVOICENO = generateInvoiceNumber(i, CStr(CLng(rawInvoiceNumber) + 1), rawLength, NAME, FORMmainDate, Doc)
                                Call replaceText(wrdDoc, "<<DATE>>", Format(FORMmainDate, "dd/mmm/yyyy"))
                                Call replaceText(wrdDoc, "<<NAME>>", NAME)
                                Call replaceText(wrdDoc, "<<ADDRESS1>>", ADDRESS1)
                                Call replaceText(wrdDoc, "<<ADDRESS2>>", ADDRESS2)
                                Call replaceText(wrdDoc, "<<INV>>", INVOICENO)
                                Call saveAndClose(NAME, ADDRESS1, Doc, FORMmainDate, monthYear)
                                On Error GoTo databaseError
                                '''ADD TO DATABASE
                                With ThisWorkbook.Sheets("Database")
                                    .Range("A2").EntireRow.Insert
                                    .Rows(2).ClearFormats
                                    .Rows(2).RowHeight = 20
                                    
                                    .Cells(2, 1).Value = NAME
                                    .Cells(2, 2).Value = ADDRESS1
                                    .Cells(2, 3).Value = ADDRESS2
                                    
                                    .Cells(2, 4).Value = NAME2
                                    .Cells(2, 5).Value = DOB
                                    .Cells(2, 6).Value = IDNO
                                    .Cells(2, 7).Value = GROUPNO
                                    
                                    .Cells(2, 8).Value = FORMmainDate
                                    .Cells(2, 9).Hyperlinks.Add .Cells(2, 9), wholeFilePath, TextToDisplay:=INVOICENO
                                    .Cells(2, 10).Value = Trim(ThisWorkbook.Sheets("Control").Range("C2").Value)
                                    .Cells(2, 13).Value = monthYear
                                    '.Cells(2, 14).Value = i
                                    '.Cells(2, 15).Value = docCounter
                                    .Range("A2:M2").Interior.Color = RGB(207, 207, 230)
                                End With
                                '''added to database
                                ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value = ThisWorkbook.Sheets(Doc).Cells(formDateIndex, 2).Value + 1
                                ThisWorkbook.Sheets(CStr(Doc)).Visible = 2
                                ThisWorkbook.Save
                            End If
                        End If
                    Next Docx
                End If
            Next i
            Call clearClassEntry
        Else
            MsgBox ("Selected range doe not have available invoice numbers. Invoice numbers could be negative or you generated too many forms for this time frame.")
        End If
    Else
        '''TODO might add highlights to bad cells
        MsgBox ("Missing Information. Cant generate forms")
    End If
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("Control").Protect ("STRIPPED_EXAMPLE_PASSWORD")
    Exit Sub
cleanUpAllOnError:
    MsgBox ("There was an error creating the invoice information for doctor : " & Doc)
    Call mainErrorCleaner
databaseError:
    MsgBox ("There was an error adding entries to the database for doctor : " & Doc)
    Call mainErrorCleaner
End Sub

Sub mainErrorCleaner()

    Call Module2.makeSuperHiddenA11
    Call Module2.makeSuperHiddenA12
    Call Module2.makeSuperHiddenB11
    Call Module2.makeSuperHiddenB21
    Call Module2.makeSuperHiddenC11
    Call Module2.makeSuperHiddenC12
    
    On Error Resume Next
    wrdDoc.Close
    
    Set wrdApp = Nothing
    Set wrdDoc = Nothing
    wholeFilePath = ""
    End

End Sub

Sub saveAndClose(pName As String, pAddress As String, Doc As String, FORMmainDate As String, monthYear As String)
    Dim pathStart As String
    Dim cleanAddress As String
    Dim pathFolder1 As String
    Dim pathFolder2 As String
    Dim pathFolder3 As String
    
    On Error GoTo saveError
    
    pathStart = Application.ActiveWorkbook.Path
    pathFolder1 = pathStart & "\MAIN WORK"
    pathFolder2 = pathFolder1 & "\" & Format(Trim(monthYear), "mmm-yyyy")
    pathFolder3 = pathFolder2 & "\" & Trim(cleanString(pName)) & " - " & Trim(cleanString(pAddress))
        
    If Len(Dir(pathFolder1, vbDirectory)) = 0 Then
       MkDir pathFolder1
    End If
    If Len(Dir(pathFolder2, vbDirectory)) = 0 Then
       MkDir pathFolder2
    End If
    If Len(Dir(pathFolder3, vbDirectory)) = 0 Then
       MkDir pathFolder3
    End If
    
    wholeFilePath = pathFolder3 & "\" & Doc & " - " & Format(FORMmainDate, "dd mmm yyyy") & ".docx"
    
    With wrdDoc
        wrdApp.DisplayAlerts = wdAlertsNone
        .SaveAs wholeFilePath, FileFormat:=12, Password:=""
        '.Close
        wrdApp.DisplayAlerts = wdAlertsAll
    End With
    
    'If bWeStartedWord Then wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    Exit Sub
saveError:
    MsgBox ("Cants save the file. Client name or address has a problem for doctor : " & Doc & ".  The invoice number was not used, database was not updated, clean exit.")
    wrdDoc.Close savechanges:=False
    Call mainErrorCleaner
End Sub

Function zeroExtend(invoiceNumber As String, rawLength As Integer) As String
    Dim stepper As Integer
    Dim extender As String
    
    extender = ""
    stepper = rawLength - Len(CStr(invoiceNumber))
    Do While stepper > 0
        extender = extender + "0"
        stepper = stepper - 1
    Loop
    zeroExtend = extender & invoiceNumber
End Function

Function generateInvoiceNumber(i As Integer, rawInvoiceNumber As String, rawLength As Integer, NAME As String, FORMmainDate As String, Doc As String) As String
    Dim invoiceStart, invoiceYear, initials As String

    invoiceYear = Right(year(CDate(FORMmainDate)), 2)

    Select Case i
    
    Case 0
        generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength)
    Case 1
        invoiceStart = "O"
        initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
        generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
    Case 2
        Select Case Doc
            Case "Back Brace - PHYSIOACTIVE"
                invoiceStart = "BA-"
                initials = Left([NAME], 1) & "/" & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Back Brace - PRO MOTION PHYSIO"
                invoiceStart = "bb-"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & "-" & initials
            Case "Back Brace - YONGE ST PHYSIO"
                invoiceStart = "BB"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Knee Brace - PHYSIOMOBILITY"
                invoiceStart = "-KB"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials & invoiceStart
            Case "Knee Brace - PRO MOTION PHYSIO"
                invoiceStart = "bk-"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & "-" & initials
            Case "Knee Brace - YONGE ST PHYSIO"
                invoiceStart = "KB"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Orthotics - BIOPED"
                invoiceStart = "C"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Orthotics - MEDIC CLINIC"
                invoiceStart = "OR"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Orthotics - PHYSIOACTIVE"
                invoiceStart = "OR-"
                initials = Left([NAME], 1) & "/" & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Chiropod - WALKIN COMFORT"
                generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength)
            Case "Orthotics - WALKIN COMFORT"
                generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength)
        End Select
    Case 3
        Select Case Doc
            Case "Back Brace - PHYSIOMED"
                invoiceStart = "BA"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Back Brace - TIMES PHYSIO"
                invoiceStart = "B-"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Knee Brace - BRAMROSE"
                invoiceStart = "KB-"
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength)
            Case "Chiropod - BRAMROSE"
                invoiceStart = "CH-"
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength)
            Case "Orthotics - BRAMROSE"
                invoiceStart = "ORT-"
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength)
        End Select
    Case 4
        Select Case Doc
            Case "Chiropractic - FOCUSED"
                invoiceStart = "C"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Massage - FOCUSED"
                invoiceStart = "M"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Physio - PHYSIOACTIVE"
                invoiceStart = "P-"
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear
            Case "Physio - PHYSIOMOBILITY"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Physio - PRO MOTION"
                invoiceStart = "P"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & "-" & initials
            Case "Physio - YONGE ST PHYSIO"
                invoiceStart = "P"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
        End Select
    Case 5
        Select Case Doc
            Case "Chiropractic - THERAPUTIX"
                invoiceStart = "C"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Massage - THERAPUTIX"
                invoiceStart = "M"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
            Case "Physio - PHYSIOMED"
                invoiceStart = "P"
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear
            Case "Physio - PROACTIVE"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = zeroExtend(rawInvoiceNumber, rawLength) & "-P" & invoiceYear & initials
            Case "Physio - TIMES PHYSIO"
                invoiceStart = "C"
                initials = Left([NAME], 1) & Mid([NAME], InStrRev([NAME], " ") + 1, 1)
                generateInvoiceNumber = invoiceStart & zeroExtend(rawInvoiceNumber, rawLength) & invoiceYear & initials
        End Select
    
    End Select
End Function

Function getFormDateIndexSpecial(Doc As String, givenDate As Date) As Integer
    Dim c_row As Integer
    c_row = 1
    
    Do While ThisWorkbook.Sheets(Doc).Cells(c_row, 1) <> "" And _
             CDate(ThisWorkbook.Sheets(Doc).Cells(c_row, 1).Value) <> givenDate
        c_row = c_row + 1
    Loop
    If ThisWorkbook.Sheets(Doc).Cells(c_row, 1) <> "" Then
        getFormDateIndexSpecial = c_row
    Else
        getFormDateIndexSpecial = -1
    End If
End Function

Function getFormDateIndex(Doc As String, formYear As String, formMonth As String) As Integer
    Dim stIndex, i, cIndex As Integer
    stIndex = getDateStIndex(Doc, formYear, formMonth)
    
    If stIndex <> -1 Then
        For i = 0 To 1
            cIndex = stIndex + i
            Do While Format(CDate(ThisWorkbook.Sheets(Doc).Cells(cIndex, 1).Value), "mmmm") = Format(CDate(ThisWorkbook.Sheets(Doc).Cells(cIndex + 2, 1).Value), "mmmm")
                If ThisWorkbook.Sheets(Doc).Cells(cIndex, 2).Value > ThisWorkbook.Sheets(Doc).Cells(cIndex + 2, 2).Value Then
                    If ThisWorkbook.Sheets(Doc).Cells(cIndex + 2, 2).Value < 25 Then
                        getFormDateIndex = cIndex + 2
                        Exit Function
                    End If
                End If
                cIndex = cIndex + 2
            Loop
        Next i
        If ThisWorkbook.Sheets(Doc).Cells(stIndex, 2).Value > ThisWorkbook.Sheets(Doc).Cells(stIndex + 1, 2).Value Then
            If ThisWorkbook.Sheets(Doc).Cells(stIndex + 1, 2).Value < 25 Then
                getFormDateIndex = stIndex + 1
                Exit Function
            End If
        Else
            If ThisWorkbook.Sheets(Doc).Cells(stIndex, 2).Value < 25 Then
                getFormDateIndex = stIndex
                Exit Function
            End If
        End If
    End If
    getFormDateIndex = -1
End Function

Function getDateStIndex(Doc As String, formYear As String, formMonth As String) As Integer
    Dim c_row As Integer
    c_row = 1
    
    Do While ThisWorkbook.Sheets(Doc).Cells(c_row, 1) <> "" And _
             CStr(Format(CDate(ThisWorkbook.Sheets(Doc).Cells(c_row, 1).Value), "mmmm yyyy")) <> formMonth & " " & formYear
        c_row = c_row + 1
    Loop
    If ThisWorkbook.Sheets(Doc).Cells(c_row, 1) <> "" Then
        getDateStIndex = c_row
    Else
        getDateStIndex = -1
    End If
End Function

Function getTitle(i As Integer) As String

    Select Case i
        Case 0
            getTitle = CStr(ThisWorkbook.Sheets("Control").Range("C2").Value) & " "
        Case Else
            getTitle = ""
    End Select

End Function

Function getDoc(i As Integer) As String
    Dim doc1 As String
    Dim doc2 As String
    Dim doc3 As String
    
    Select Case i
        Case 0
            getDoc = CStr(ThisWorkbook.Sheets("Control").Range("H2").Value)
            getDoc = Mid(getDoc, 4, Len(getDoc) - 3)
        Case 1
            getDoc = CStr(ThisWorkbook.Sheets("Control").Range("H4").Value)
            getDoc = Mid(getDoc, 4, Len(getDoc) - 3)
        Case 2
            doc1 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H7").Value))
            If Len(doc1) > 7 Then
                doc1 = Mid(doc1, 4, Len(doc1) - 3)
            Else
                doc1 = ""
            End If
            
            doc2 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H8").Value))
            If Len(doc2) > 7 Then
                doc2 = Mid(doc2, 4, Len(doc2) - 3)
            Else
                doc2 = ""
            End If
            
            If (doc2 = "Chiropod / Orthotics - WALKIN COMFORT") Then
                doc2 = "Chiropod - WALKIN COMFORT"
            End If
            
            getDoc = doc1 + "|" + doc2
        Case 3
            doc1 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H10").Value))
            If Len(doc1) > 7 Then
                doc1 = Mid(doc1, 4, Len(doc1) - 3)
            Else
                doc1 = ""
            End If
            
            doc2 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H11").Value))
            If Len(doc2) > 7 Then
                doc2 = Mid(doc2, 4, Len(doc2) - 3)
            Else
                doc2 = ""
            End If
            
            If (doc2 = "Chiropod / Orthotics - BRAMROSE") Then
                doc2 = "Chiropod - BRAMROSE"
            End If
            
            getDoc = doc1 + "|" + doc2
        Case 4
            doc1 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H14").Value))
            If Len(doc1) > 7 Then
                doc1 = Mid(doc1, 4, Len(doc1) - 3)
            Else
                doc1 = ""
            End If
            
            doc2 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H15").Value))
            If Len(doc2) > 7 Then
                doc2 = Mid(doc2, 4, Len(doc2) - 3)
            Else
                doc2 = ""
            End If
            
            doc3 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H16").Value))
            If Len(doc3) > 7 Then
                doc3 = Mid(doc3, 4, Len(doc3) - 3)
            Else
                doc3 = ""
            End If
            
            getDoc = doc1 + "|" + doc2 + "|" + doc3
        Case 5
            doc1 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H19").Value))
            If Len(doc1) > 7 Then
                doc1 = Mid(doc1, 4, Len(doc1) - 3)
            Else
                doc1 = ""
            End If
            
            doc2 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H20").Value))
            If Len(doc2) > 7 Then
                doc2 = Mid(doc2, 4, Len(doc2) - 3)
            Else
                doc2 = ""
            End If
            
            doc3 = CStr(Trim(ThisWorkbook.Sheets("Control").Range("H21").Value))
            If Len(doc3) > 7 Then
                doc3 = Mid(doc3, 4, Len(doc3) - 3)
            Else
                doc3 = ""
            End If
            
            getDoc = doc1 + "|" + doc2 + "|" + doc3
    End Select

End Function

Function checkSelectedFormVals() As Boolean

    checkSelectedFormVals = (Not selectedForms(0) Or checkForm11Vals) And _
                            (Not selectedForms(1) Or checkForm12Vals) And _
                            (Not selectedForms(2) Or checkForm21Vals) And _
                            (Not selectedForms(3) Or checkForm22Vals) And _
                            (Not selectedForms(4) Or checkForm31Vals) And _
                            (Not selectedForms(5) Or checkForm32Vals)

End Function

Function isFilled(inputString As String) As Boolean
    isFilled = Trim(inputString) <> ""
End Function


Function checkForm11Vals() As Boolean
    With ThisWorkbook.Sheets("Control")
        checkForm11Vals = baseCheck() And _
                          isFilled(.Range("H2")) And _
                          isFilled(.Range("C2")) And _
                          isFilled(.Range("I2")) And _
                          isFilled(.Range("J2"))
    End With
End Function

Function checkForm12Vals() As Boolean
    With ThisWorkbook.Sheets("Control")
        checkForm12Vals = baseCheck() And baseCheck2() And _
                          isFilled(.Range("H4")) And _
                          isFilled(.Range("I4")) And _
                          isFilled(.Range("J4"))
    End With
End Function

Function checkForm21Vals() As Boolean
    With ThisWorkbook.Worksheets("Control")
        checkForm21Vals = baseCheck() And _
                          ((isFilled(.Range("H7")) And isFilled(.Range("I7")) And isFilled(.Range("J7"))) Or (Not isFilled(.Range("H7")))) And _
                          ((isFilled(.Range("H8")) And isFilled(.Range("I8")) And isFilled(.Range("J8"))) Or (Not isFilled(.Range("H8"))))
    End With
End Function

Function checkForm22Vals() As Boolean
    With ThisWorkbook.Worksheets("Control")
        checkForm22Vals = baseCheck() And baseCheck2() And _
                          ((isFilled(.Range("H10")) And isFilled(.Range("I10")) And isFilled(.Range("J10"))) Or (Not isFilled(.Range("H10")))) And _
                          ((isFilled(.Range("H11")) And isFilled(.Range("I11")) And isFilled(.Range("J11"))) Or (Not isFilled(.Range("H11"))))
    End With
End Function

Function checkForm31Vals() As Boolean
    With ThisWorkbook.Worksheets("Control")
        checkForm31Vals = baseCheck() And _
                          ((isFilled(.Range("H14")) And isFilled(.Range("K14")) And isFilled(.Range("L14")) And isFilled(.Range("M14"))) Or (Not isFilled(.Range("H14")))) And _
                          ((isFilled(.Range("H15")) And isFilled(.Range("K15")) And isFilled(.Range("L15")) And isFilled(.Range("M15"))) Or (Not isFilled(.Range("H15")))) And _
                          ((isFilled(.Range("H16")) And isFilled(.Range("K16")) And isFilled(.Range("L16")) And isFilled(.Range("M16"))) Or (Not isFilled(.Range("H16"))))
    End With
End Function

Function checkForm32Vals() As Boolean
    With ThisWorkbook.Worksheets("Control")
        checkForm32Vals = baseCheck() And baseCheck2() And _
                          ((isFilled(.Range("H19")) And isFilled(.Range("K19")) And isFilled(.Range("L19")) And isFilled(.Range("M19"))) Or (Not isFilled(.Range("H19")))) And _
                          ((isFilled(.Range("H20")) And isFilled(.Range("K20")) And isFilled(.Range("L20")) And isFilled(.Range("M20"))) Or (Not isFilled(.Range("H20")))) And _
                          ((isFilled(.Range("H21")) And isFilled(.Range("K21")) And isFilled(.Range("L21")) And isFilled(.Range("M21"))) Or (Not isFilled(.Range("H21"))))
    End With
End Function

Function baseCheck() As Boolean
    baseCheck = Trim(ThisWorkbook.Worksheets("Control").Range("C3").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C4").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C5").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C6").Value) <> ""
End Function

Function baseCheck2() As Boolean
    baseCheck2 = Trim(ThisWorkbook.Worksheets("Control").Range("C7").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C8").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C9").Value) <> "" And _
                      Trim(ThisWorkbook.Worksheets("Control").Range("C10").Value) <> ""
End Function

Function yearMonthCheck() As Boolean
    yearMonthCheck = Trim(ThisWorkbook.Sheets("Control").Range("C12")) <> "" And _
                     Trim(ThisWorkbook.Sheets("Control").Range("C13")) <> ""
End Function

Function cleanString(pAddress As String) As String 'Cants have folders and files with /\:*?"<>|
    pAddress = Replace(pAddress, "/", "")
    pAddress = Replace(pAddress, "\", "")
    pAddress = Replace(pAddress, ":", "")
    pAddress = Replace(pAddress, "*", "")
    pAddress = Replace(pAddress, "?", "")
    pAddress = Replace(pAddress, """", "")
    pAddress = Replace(pAddress, "<", "")
    pAddress = Replace(pAddress, ">", "")
    pAddress = Replace(pAddress, "|", "")
    
    If InStr(pAddress, ",") <> 0 Then
        cleanString = Left(pAddress, InStr(pAddress, ",") - 1)
    Else
        cleanString = pAddress
    End If

End Function

Function areThereAvaliableDates() As Boolean
        
        'Dim bookYear As String
        'Dim bookMonth As String
        
        'bookYear = ThisWorkbook.Sheets("Control").Range("C12")
        'bookMonth = ThisWorkbook.Sheets("Control").Range("C13")
        Dim q0, q1, q2, q3, q4, q5 As Boolean
        Dim docs As String
        q0 = True
        q1 = True
        q2 = True
        q3 = True
        q4 = True
        q5 = True
        If selectedForms(0) Then
            q0 = getFormDateIndex(getDoc(0), ThisWorkbook.Sheets("Control").Range("I2"), ThisWorkbook.Sheets("Control").Range("J2")) <> -1
        End If
        If selectedForms(1) Then
            q1 = getFormDateIndex(getDoc(1), ThisWorkbook.Sheets("Control").Range("I4"), ThisWorkbook.Sheets("Control").Range("J4")) <> -1
        End If
        If selectedForms(2) Then
            docs = getDoc(2)
            If Split(docs, "|")(0) = "" Then
                q2 = True
            Else
                q2 = getFormDateIndex(CStr(Split(docs, "|")(0)), ThisWorkbook.Sheets("Control").Range("I7"), ThisWorkbook.Sheets("Control").Range("J7")) <> -1
            End If
            
            If Split(docs, "|")(1) = "" Then
                q3 = True
            Else
                q3 = getFormDateIndex(CStr(Split(docs, "|")(1)), ThisWorkbook.Sheets("Control").Range("I8"), ThisWorkbook.Sheets("Control").Range("J8")) <> -1
            End If
        End If
        If selectedForms(3) Then
            docs = getDoc(3)
            If Split(docs, "|")(0) = "" Then
                q4 = True
            Else
                q4 = getFormDateIndex(CStr(Split(docs, "|")(0)), ThisWorkbook.Sheets("Control").Range("I10"), ThisWorkbook.Sheets("Control").Range("J10")) <> -1
            End If
            
            If Split(docs, "|")(1) = "" Then
                q5 = True
            Else
                q5 = getFormDateIndex(CStr(Split(docs, "|")(1)), ThisWorkbook.Sheets("Control").Range("I11"), ThisWorkbook.Sheets("Control").Range("J11")) <> -1
            End If
        End If
        areThereAvaliableDates = q0 And q1 And q2 And q3 And q4 And q5

End Function

Sub clearAllNew()
    With ThisWorkbook.Sheets("Control")
        .Unprotect ("STRIPPED_EXAMPLE_PASSWORD")
        .Range("C2:C10") = ""
        .Range("H2:H21") = ""
        .Range("K14:M16") = ""
        .Range("K19:M21") = ""
        .Range("I2:J2") = ""
        .Range("I4:J4") = ""
        .Range("I7:J8") = ""
        .Range("I10:J11") = ""
        For i = 0 To 5
            .Shapes("Check Box " & (i + 1)).OLEFormat.Object.Value = 0
        Next i
        .Protect ("STRIPPED_EXAMPLE_PASSWORD")
    End With
End Sub

Sub clearClassEntry()
    With ThisWorkbook.Sheets("Control")
        .Unprotect ("STRIPPED_EXAMPLE_PASSWORD")
        .Range("H2:H21") = ""
        .Range("K14:M16") = ""
        .Range("K19:M21") = ""
        .Range("I2:J2") = ""
        .Range("I4:J4") = ""
        .Range("I7:J8") = ""
        .Range("I10:J11") = ""
        For i = 0 To 5
            .Shapes("Check Box " & (i + 1)).OLEFormat.Object.Value = 0
        Next i
        .Protect ("STRIPPED_EXAMPLE_PASSWORD")
    End With
End Sub


