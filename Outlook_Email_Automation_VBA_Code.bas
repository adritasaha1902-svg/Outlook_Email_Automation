Attribute VB_Name = "Module2"
Option Explicit

Public Sub Run_Department_Email_Automation()

    Dim wbTarget As Workbook
    Dim wsEmail As Worksheet
    Dim wsControl As Worksheet

    Dim OutlookApp As Object
    Dim OutlookMail As Object

    Dim lastRow As Long
    Dim i As Long

    Dim department As String
    Dim recipientName As String
    Dim recipientEmail As String
    Dim reportSheetName As String
    Dim activeFlag As String
    Dim emailMode As String
    Dim reportingPeriod As String
    Dim senderName As String
    Dim companyName As String

    Dim reportFolder As String
    Dim reportPath As String
    Dim emailHTML As String
    Dim statusText As String
    Dim errorText As String

    Dim processedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    On Error GoTo FatalError

    ' Preserve Excel settings
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    ' Reduce Excel processing overhead
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbTarget = GetAutomationWorkbook()

    If wbTarget Is Nothing Then
        Err.Raise vbObjectError + 200, , _
            "A workbook containing both 'Email_List' and " & _
            "'Email_Control_List' could not be found."
    End If

    Set wsEmail = wbTarget.Worksheets("Email_List")
    Set wsControl = wbTarget.Worksheets("Email_Control_List")

    emailMode = UCase$(Trim$(CStr(wsControl.Range("B2").value)))
    reportingPeriod = Trim$(CStr(wsControl.Range("B3").value))
    senderName = Trim$(CStr(wsControl.Range("B4").value))
    companyName = Trim$(CStr(wsControl.Range("B5").value))

    If emailMode <> "SEND" And emailMode <> "DISPLAY" Then
        Err.Raise vbObjectError + 201, , _
            "Enter SEND or DISPLAY in Email_Control_List!B2."
    End If

    If Len(wbTarget.Path) = 0 Then
        Err.Raise vbObjectError + 202, , _
            "Save the target workbook before running the macro."
    End If

    If emailMode = "SEND" Then

        If MsgBox( _
            "You are about to send department reports through Outlook." & _
            vbCrLf & vbCrLf & _
            "Continue?", _
            vbYesNo + vbQuestion, _
            "Confirm Email Sending") <> vbYes Then

            GoTo CleanExit

        End If

    End If

    reportFolder = wbTarget.Path & Application.PathSeparator & _
                   "Department_Reports"

    If Dir(reportFolder, vbDirectory) = vbNullString Then
        MkDir reportFolder
    End If

    ' Reuse an existing Classic Outlook instance when possible
    On Error Resume Next

    Set OutlookApp = GetObject(, "Outlook.Application")

    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If

    On Error GoTo FatalError

    If OutlookApp Is Nothing Then
        Err.Raise vbObjectError + 203, , _
            "Unable to connect to Classic Outlook."
    End If

    lastRow = wsEmail.Cells(wsEmail.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow

        department = Trim$(CStr(wsEmail.Cells(i, "A").value))
        recipientName = Trim$(CStr(wsEmail.Cells(i, "B").value))
        recipientEmail = Trim$(CStr(wsEmail.Cells(i, "C").value))
        reportSheetName = Trim$(CStr(wsEmail.Cells(i, "D").value))
        activeFlag = UCase$(Trim$(CStr(wsEmail.Cells(i, "E").value)))

        reportPath = vbNullString
        statusText = "Failed"
        errorText = vbNullString

        On Error GoTo RecipientError

        If activeFlag <> "Y" And activeFlag <> "YES" Then
            GoTo NextRecipient
        End If

        If department = vbNullString Then
            Err.Raise vbObjectError + 204, , _
                "Department is blank."
        End If

        If recipientName = vbNullString Then
            Err.Raise vbObjectError + 205, , _
                "Recipient name is blank."
        End If

        If Not IsValidEmail(recipientEmail) Then
            Err.Raise vbObjectError + 206, , _
                "Invalid email address: " & recipientEmail
        End If

        If reportSheetName = vbNullString Then
            reportSheetName = GetDepartmentSheetName(department)
        End If

        If Not WorksheetExistsInWorkbook(wbTarget, reportSheetName) Then
            Err.Raise vbObjectError + 207, , _
                "Report worksheet not found: " & reportSheetName
        End If

        If IsDuplicateEmail( _
            wbTarget, _
            department, _
            recipientEmail, _
            reportingPeriod, _
            emailMode) Then

            statusText = "Duplicate Skipped"
            skippedCount = skippedCount + 1

            WriteEmailLog _
                wbTarget, _
                department, _
                recipientName, _
                recipientEmail, _
                reportingPeriod, _
                reportPath, _
                statusText, _
                vbNullString

            GoTo NextRecipient

        End If

        ' Generate the department PDF
        reportPath = GenerateDepartmentReport( _
                        wbTarget, _
                        reportSheetName, _
                        department, _
                        reportingPeriod, _
                        reportFolder)

        DoEvents

        If Len(Dir(reportPath)) = 0 Then
            Err.Raise vbObjectError + 208, , _
                "The PDF report could not be located: " & reportPath
        End If

        ' Build email body with the department summary table
        emailHTML = BuildHTMLEmail( _
                        recipientName, _
                        department, _
                        reportingPeriod, _
                        senderName, _
                        companyName, _
                        wbTarget, _
                        reportSheetName)

        DoEvents

        Set OutlookMail = OutlookApp.CreateItem(0)

        If OutlookMail Is Nothing Then
            Err.Raise vbObjectError + 209, , _
                "Unable to create the Outlook email item."
        End If

        With OutlookMail

            .To = recipientEmail

            .Subject = department & _
                       " Department Report - " & _
                       reportingPeriod

            .HTMLBody = emailHTML

            .ReadReceiptRequested = False
            .OriginatorDeliveryReportRequested = False

            DoEvents

            .Attachments.Add reportPath

            DoEvents

            If emailMode = "SEND" Then

                .Send
                statusText = "Sent (submitted to Outlook)"

            Else

                .Display
                statusText = "Displayed"

            End If

        End With

        processedCount = processedCount + 1

        WriteEmailLog _
            wbTarget, _
            department, _
            recipientName, _
            recipientEmail, _
            reportingPeriod, _
            reportPath, _
            statusText, _
            vbNullString

        Set OutlookMail = Nothing

        DoEvents

        GoTo NextRecipient

RecipientError:

        errorText = Err.Description
        statusText = "Failed"
        failedCount = failedCount + 1

        WriteEmailLog _
            wbTarget, _
            department, _
            recipientName, _
            recipientEmail, _
            reportingPeriod, _
            reportPath, _
            statusText, _
            errorText

        Err.Clear

        If Not OutlookMail Is Nothing Then
            On Error Resume Next
            OutlookMail.Close 1
            On Error GoTo 0
        End If

        Set OutlookMail = Nothing

        On Error GoTo FatalError

NextRecipient:

        DoEvents
        On Error GoTo FatalError

    Next i

    MsgBox _
        "Department Email Automation Completed!" & _
        vbCrLf & vbCrLf & _
        "Emails processed: " & processedCount & _
        vbCrLf & _
        "Duplicates skipped: " & skippedCount & _
        vbCrLf & _
        "Failed: " & failedCount, _
        vbInformation, _
        "Automation Complete"

CleanExit:

    On Error Resume Next

    If Not OutlookMail Is Nothing Then
        OutlookMail.Close 1
    End If

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation

    On Error GoTo 0

    Exit Sub

FatalError:

    MsgBox _
        "Automation error: " & Err.Description, _
        vbCritical, _
        "Automation Error"

    Resume CleanExit

End Sub


Private Function GetAutomationWorkbook() As Workbook

    Dim wb As Workbook

    If Not ActiveWorkbook Is Nothing Then

        If WorksheetExistsInWorkbook(ActiveWorkbook, "Email_List") _
           And WorksheetExistsInWorkbook(ActiveWorkbook, "Email_Control_List") Then

            Set GetAutomationWorkbook = ActiveWorkbook
            Exit Function

        End If

    End If

    If Not ThisWorkbook Is Nothing Then

        If WorksheetExistsInWorkbook(ThisWorkbook, "Email_List") _
           And WorksheetExistsInWorkbook(ThisWorkbook, "Email_Control_List") Then

            Set GetAutomationWorkbook = ThisWorkbook
            Exit Function

        End If

    End If

    For Each wb In Application.Workbooks

        If WorksheetExistsInWorkbook(wb, "Email_List") _
           And WorksheetExistsInWorkbook(wb, "Email_Control_List") Then

            Set GetAutomationWorkbook = wb
            Exit Function

        End If

    Next wb

    Set GetAutomationWorkbook = Nothing

End Function


Private Function WorksheetExistsInWorkbook( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    WorksheetExistsInWorkbook = False

    If wb Is Nothing Then Exit Function

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    On Error GoTo 0

    WorksheetExistsInWorkbook = Not ws Is Nothing

End Function


Private Function GetDepartmentSheetName( _
    ByVal department As String) As String

    department = Trim$(department)

    If LCase$(Right$(department, 7)) = "_report" Then
        GetDepartmentSheetName = department
    Else
        GetDepartmentSheetName = department & "_Report"
    End If

End Function


Private Function GenerateDepartmentReport( _
    ByVal wb As Workbook, _
    ByVal reportSheetName As String, _
    ByVal department As String, _
    ByVal reportingPeriod As String, _
    ByVal reportFolder As String) As String

    Dim wsSource As Worksheet
    Dim pdfPath As String
    Dim safeDepartment As String
    Dim safePeriod As String

    On Error GoTo ReportError

    Set wsSource = wb.Worksheets(reportSheetName)

    safeDepartment = CleanFileName(department)
    safePeriod = CleanFileName(reportingPeriod)

    pdfPath = reportFolder & Application.PathSeparator & _
              safeDepartment & "_" & _
              safePeriod & "_Report.pdf"

    If Len(Dir(pdfPath)) > 0 Then
        Kill pdfPath
    End If

    DoEvents

    wsSource.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    If Not WaitForFile(pdfPath, 10) Then
        Err.Raise vbObjectError + 210, , _
            "PDF was not created within the expected time: " & pdfPath
    End If

    GenerateDepartmentReport = pdfPath

    Exit Function

ReportError:

    Err.Raise vbObjectError + 211, , _
        "Unable to generate PDF for " & _
        department & ": " & Err.Description

End Function


Private Function WaitForFile( _
    ByVal filePath As String, _
    ByVal timeoutSeconds As Long) As Boolean

    Dim startTime As Single
    Dim elapsedTime As Single

    startTime = Timer

    Do While Len(Dir(filePath)) = 0

        DoEvents

        elapsedTime = Timer - startTime

        If elapsedTime < 0 Then
            elapsedTime = elapsedTime + 86400
        End If

        If elapsedTime >= timeoutSeconds Then
            WaitForFile = False
            Exit Function
        End If

    Loop

    WaitForFile = True

End Function


Private Function BuildHTMLEmail( _
    ByVal recipientName As String, _
    ByVal department As String, _
    ByVal reportingPeriod As String, _
    ByVal senderName As String, _
    ByVal companyName As String, _
    ByVal wb As Workbook, _
    ByVal reportSheetName As String) As String

    Dim html As String
    Dim reportSummaryHTML As String

    reportSummaryHTML = BuildReportSummaryHTML( _
                            wb, _
                            reportSheetName, _
                            department)

    html = "<html><body style='font-family:Arial;" & _
           "font-size:10pt;color:#333;'>"

    html = html & _
           "<div style='background-color:#1F4E79;" & _
           "padding:15px;color:white;'>" & _
           "<h2 style='margin:0;'>Department Performance Report</h2>" & _
           "<p style='margin:5px 0 0 0;'>" & _
           HtmlEncode(companyName) & _
           "</p></div>"

    html = html & _
           "<p>Dear " & _
           HtmlEncode(recipientName) & _
           ",</p>"

    html = html & _
           "<p>Please find attached the <b>" & _
           HtmlEncode(department) & _
           " Department Report</b> for <b>" & _
           HtmlEncode(reportingPeriod) & _
           "</b>.</p>"

    html = html & _
           "<p>Below is a summary of the data included in the attached report:</p>"

    html = html & reportSummaryHTML

    html = html & _
           "<p>Please review the attached PDF report and share your feedback.</p>"

    html = html & _
           "<p>Regards,<br><b>" & _
           HtmlEncode(senderName) & _
           "</b><br>" & _
           HtmlEncode(companyName) & _
           "</p>"

    html = html & "</body></html>"

    BuildHTMLEmail = html

End Function


Private Function BuildReportSummaryHTML( _
    ByVal wb As Workbook, _
    ByVal reportSheetName As String, _
    ByVal department As String) As String

    Dim ws As Worksheet
    Dim usedRng As Range

    Dim html As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim displayRows As Long

    Dim r As Long
    Dim c As Long

    Dim headerText As String
    Dim cellText As String
    Dim cellValue As Variant

    Dim numericTotal As Double
    Dim numericCount As Long
    Dim averageValue As Double
    Dim hasNumericData As Boolean

    On Error GoTo SummaryError

    Set ws = wb.Worksheets(reportSheetName)
    Set usedRng = ws.UsedRange

    rowCount = usedRng.Rows.Count
    colCount = usedRng.Columns.Count

    If rowCount < 2 Or colCount < 1 Then

        BuildReportSummaryHTML = _
            "<p><i>No report data was available for this department.</i></p>"

        Exit Function

    End If

    html = _
        "<h3 style='color:#1F4E79;'>Report Summary - " & _
        HtmlEncode(department) & _
        "</h3>"

    html = html & _
        "<p><b>Data rows:</b> " & _
        Format$(rowCount - 1, "#,##0") & _
        "</p>"

    html = html & _
        "<table border='1' cellpadding='5' cellspacing='0' " & _
        "style='border-collapse:collapse;font-family:Arial;" & _
        "font-size:9pt;width:100%;'>"

    html = html & _
        "<tr style='background-color:#D9EAF7;font-weight:bold;'>" & _
        "<th align='left'>Metric</th>" & _
        "<th align='right'>Total</th>" & _
        "<th align='right'>Average</th>" & _
        "</tr>"

    hasNumericData = False

    For c = 1 To colCount

        headerText = SafeCellText(usedRng.Cells(1, c))

        If headerText = vbNullString Then
            headerText = "Column " & CStr(c)
        End If

        numericTotal = 0
        numericCount = 0

        For r = 2 To rowCount

            cellValue = usedRng.Cells(r, c).Value2

            If Not IsError(cellValue) Then

                If Len(Trim$(CStr(cellValue))) > 0 Then

                    If IsNumeric(cellValue) Then
                        numericTotal = numericTotal + CDbl(cellValue)
                        numericCount = numericCount + 1
                    End If

                End If

            End If

        Next r

        If numericCount > 0 Then

            hasNumericData = True
            averageValue = numericTotal / numericCount

            html = html & _
                "<tr>" & _
                "<td>" & HtmlEncode(headerText) & "</td>" & _
                "<td align='right'>" & _
                FormatSummaryNumber(headerText, numericTotal) & _
                "</td>" & _
                "<td align='right'>" & _
                FormatSummaryNumber(headerText, averageValue) & _
                "</td>" & _
                "</tr>"

        End If

    Next c

    If Not hasNumericData Then

        html = html & _
            "<tr><td colspan='3'>" & _
            "<i>No numeric metrics were detected.</i>" & _
            "</td></tr>"

    End If

    html = html & "</table>"

    html = html & _
        "<h4 style='color:#1F4E79;'>Report Preview</h4>"

    html = html & _
        "<table border='1' cellpadding='5' cellspacing='0' " & _
        "style='border-collapse:collapse;font-family:Arial;" & _
        "font-size:9pt;width:100%;'>"

    html = html & _
        "<tr style='background-color:#1F4E79;color:white;font-weight:bold;'>"

    For c = 1 To colCount

        html = html & _
            "<th align='left'>" & _
            HtmlEncode(SafeCellText(usedRng.Cells(1, c))) & _
            "</th>"

    Next c

    html = html & "</tr>"

    displayRows = rowCount - 1

    If displayRows > 8 Then
        displayRows = 8
    End If

    For r = 2 To displayRows + 1

        html = html & "<tr>"

        For c = 1 To colCount

            cellText = SafeCellText(usedRng.Cells(r, c))

            html = html & _
                "<td>" & _
                HtmlEncode(cellText) & _
                "</td>"

        Next c

        html = html & "</tr>"

    Next r

    If rowCount - 1 > displayRows Then

        html = html & _
            "<tr><td colspan='" & _
            CStr(colCount) & _
            "' style='font-style:italic;color:#666;'>" & _
            "Preview limited to the first " & _
            CStr(displayRows) & _
            " rows. See the attached PDF for the complete report." & _
            "</td></tr>"

    End If

    html = html & "</table>"

    BuildReportSummaryHTML = html

    Exit Function

SummaryError:

    BuildReportSummaryHTML = _
        "<p style='color:#C00000;'>" & _
        "Unable to build the email summary table: " & _
        HtmlEncode(Err.Description) & _
        "</p>"

End Function


Private Function SafeCellText( _
    ByVal targetCell As Range) As String

    If IsError(targetCell.value) Then
        SafeCellText = CStr(targetCell.Text)
    Else
        SafeCellText = Trim$(CStr(targetCell.Text))
    End If

End Function


Private Function FormatSummaryNumber( _
    ByVal headerText As String, _
    ByVal numericValue As Double) As String

    Dim headerLower As String

    headerLower = LCase$(Trim$(headerText))

    If InStr(1, headerLower, "%", vbTextCompare) > 0 _
       Or InStr(1, headerLower, "percent", vbTextCompare) > 0 _
       Or InStr(1, headerLower, "attendance", vbTextCompare) > 0 _
       Or InStr(1, headerLower, "margin", vbTextCompare) > 0 _
       Or InStr(1, headerLower, "rate", vbTextCompare) > 0 Then

        If Abs(numericValue) <= 1 Then
            FormatSummaryNumber = Format$(numericValue, "0.0%")
        Else
            FormatSummaryNumber = Format$(numericValue, "0.0")
        End If

    Else

        FormatSummaryNumber = Format$(numericValue, "#,##0.00")

    End If

End Function


Private Function CleanFileName( _
    ByVal value As String) As String

    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array( _
        "\", "/", ":", "*", "?", _
        """", "<", ">", "|")

    For Each item In invalidChars
        value = Replace(value, CStr(item), "_")
    Next item

    CleanFileName = Trim$(value)

End Function


Private Function HtmlEncode( _
    ByVal value As String) As String

    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    value = Replace(value, "'", "&#39;")

    HtmlEncode = value

End Function


Private Function IsValidEmail( _
    ByVal emailAddress As String) As Boolean

    Dim atPosition As Long
    Dim dotPosition As Long

    emailAddress = Trim$(emailAddress)

    atPosition = InStr(1, emailAddress, "@", vbTextCompare)
    dotPosition = InStrRev(emailAddress, ".")

    IsValidEmail = _
        (atPosition > 1 And _
         dotPosition > atPosition + 1 And _
         dotPosition < Len(emailAddress))

End Function


Private Function IsDuplicateEmail( _
    ByVal wb As Workbook, _
    ByVal department As String, _
    ByVal recipientEmail As String, _
    ByVal reportingPeriod As String, _
    ByVal emailMode As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim previousStatus As String

    Set ws = GetOrCreateLogSheet(wb)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow

        If StrComp( _
            Trim$(CStr(ws.Cells(r, 3).value)), _
            department, _
            vbTextCompare) = 0 Then

            If StrComp( _
                Trim$(CStr(ws.Cells(r, 5).value)), _
                recipientEmail, _
                vbTextCompare) = 0 Then

                If StrComp( _
                    Trim$(CStr(ws.Cells(r, 6).value)), _
                    reportingPeriod, _
                    vbTextCompare) = 0 Then

                    previousStatus = _
                        Trim$(CStr(ws.Cells(r, 8).value))

                    If emailMode = "SEND" Then

                        If LCase$(previousStatus) Like "sent*" Then
                            IsDuplicateEmail = True
                            Exit Function
                        End If

                    Else

                        If previousStatus = "Displayed" _
                           Or LCase$(previousStatus) Like "sent*" Then

                            IsDuplicateEmail = True
                            Exit Function

                        End If

                    End If

                End If

            End If

        End If

    Next r

    IsDuplicateEmail = False

End Function


Private Function GetOrCreateLogSheet( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets("Email_Log")

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = wb.Worksheets.Add( _
                    After:=wb.Worksheets(wb.Worksheets.Count))

        ws.Name = "Email_Log"

        ws.Range("A1:I1").value = Array( _
            "Date", _
            "Time", _
            "Department", _
            "Recipient Name", _
            "Email ID", _
            "Reporting Period", _
            "Report Path", _
            "Status", _
            "Error Details")

        ws.Rows(1).Font.Bold = True
        ws.Columns("A:I").AutoFit

    End If

    Set GetOrCreateLogSheet = ws

End Function


Private Sub WriteEmailLog( _
    ByVal wb As Workbook, _
    ByVal department As String, _
    ByVal recipientName As String, _
    ByVal recipientEmail As String, _
    ByVal reportingPeriod As String, _
    ByVal reportPath As String, _
    ByVal statusText As String, _
    ByVal errorDetails As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo LogError

    Set ws = GetOrCreateLogSheet(wb)

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, 1).value = Date
    ws.Cells(nextRow, 2).value = Time
    ws.Cells(nextRow, 3).value = department
    ws.Cells(nextRow, 4).value = recipientName
    ws.Cells(nextRow, 5).value = recipientEmail
    ws.Cells(nextRow, 6).value = reportingPeriod
    ws.Cells(nextRow, 7).value = reportPath
    ws.Cells(nextRow, 8).value = statusText
    ws.Cells(nextRow, 9).value = errorDetails

    Exit Sub

LogError:

    Debug.Print "Email log error: " & Err.Description

End Sub

