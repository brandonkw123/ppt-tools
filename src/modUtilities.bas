Attribute VB_Name = "modUtilities"
' ============================================================
' modUtilities — Email and export commands
' ============================================================

Public Sub EmailSelectedSlides(control As IRibbonControl)
    ' Exports selected slides (from the slide panel) to a new .pptx file
    ' and opens a new Outlook email with it attached

    If ActiveWindow.Selection.Type <> ppSelectionSlides Then
        MsgBox "Please select one or more slides in the slide panel.", vbExclamation
        Exit Sub
    End If

    Dim sourcePres As Presentation
    Dim newPres As Presentation
    Dim tempPath As String
    Dim sr As SlideRange

    Set sourcePres = ActivePresentation
    Set sr = ActiveWindow.Selection.SlideRange

    ' Create a new blank presentation to copy slides into
    Set newPres = Presentations.Add(WithWindow:=msoFalse)

    ' Copy each selected slide into the new presentation
    Dim i As Integer
    For i = 1 To sr.Count
        sr(i).Copy
        newPres.Slides.Paste
    Next i

    ' Remove the blank first slide that Presentations.Add creates
    newPres.Slides(1).Delete

    ' Save to temp location
    tempPath = Environ("TEMP") & "\Selected Slides.pptx"
    newPres.SaveAs tempPath, ppSaveAsOpenXMLPresentation
    newPres.Close

    ' Open Outlook and attach the file
    Dim olApp As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0) ' 0 = olMailItem

    With olMail
        .Subject = "Selected Slides"
        .Attachments.Add tempPath
        .Display ' Opens the email window for the user to address and send
    End With
End Sub

Public Sub EmailWholeDeck(control As IRibbonControl)
    ' Saves the current presentation to a temp location and opens
    ' a new Outlook email with it attached

    Dim pres As Presentation
    Set pres = ActivePresentation

    If pres.Path = "" Then
        MsgBox "Please save your presentation before emailing it.", vbExclamation
        Exit Sub
    End If

    ' Save current state before attaching
    pres.Save

    Dim olApp As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .Subject = pres.Name
        .Attachments.Add pres.FullName
        .Display
    End With
End Sub

Public Sub ConvertToPDF(control As IRibbonControl)
    ' One-click export to PDF, saved in the same folder as the presentation

    Dim pres As Presentation
    Set pres = ActivePresentation

    If pres.Path = "" Then
        MsgBox "Please save your presentation before converting to PDF.", vbExclamation
        Exit Sub
    End If

    Dim pdfPath As String
    pdfPath = pres.Path & "\" & Replace(pres.Name, ".pptx", ".pdf")
    pdfPath = Replace(pdfPath, ".pptm", ".pdf")

    pres.ExportAsFixedFormat pdfPath, ppFixedFormatTypePDF

    MsgBox "PDF saved to:" & vbNewLine & pdfPath, vbInformation
End Sub
