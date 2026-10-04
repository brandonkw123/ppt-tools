Attribute VB_Name = "modUtilities"
' ============================================================
' modUtilities - Email and export commands
' ============================================================

Public Sub EmailSelectedSlides(control As IRibbonControl)
    ' Saves a copy of the deck containing only the slides selected in the slide panel
    ' and opens a new Outlook email with it attached

    If Application.Windows.Count = 0 Then
        Beep
        Exit Sub
    End If
    If ActiveWindow.Selection.Type <> ppSelectionSlides Then
        Beep
        Exit Sub
    End If

    Dim sourcePres As Presentation
    Dim newPres As Presentation
    Dim tempPath As String
    Dim sr As SlideRange

    Set sourcePres = ActivePresentation
    Set sr = ActiveWindow.Selection.SlideRange

    ' Remember the selected slides by ID (IDs survive the copy; indexes shift as slides are deleted)
    Dim keepIDs As String
    Dim i As Long
    keepIDs = "|"
    For i = 1 To sr.Count
        keepIDs = keepIDs & sr(i).SlideID & "|"
    Next i

    ' Save a full copy so the original theme, layouts and formatting are kept
    tempPath = Environ("TEMP") & "\Selected Slides.pptx"
    sourcePres.SaveCopyAs tempPath, ppSaveAsOpenXMLPresentation

    ' Open the copy without a window and delete every slide that wasn't selected
    Set newPres = Presentations.Open(tempPath, WithWindow:=msoFalse)
    For i = newPres.Slides.Count To 1 Step -1
        If InStr(keepIDs, "|" & newPres.Slides(i).SlideID & "|") = 0 Then
            newPres.Slides(i).Delete
        End If
    Next i
    newPres.Save
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
    ' Saves the current presentation and opens
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
    ' Export to PDF with a save dialog for the user to choose location

    Dim pres As Presentation
    Set pres = ActivePresentation

    ' Default file name: the deck's name with its extension swapped for .pdf
    Dim defaultName As String
    defaultName = pres.Name
    If InStrRev(defaultName, ".") > 0 Then defaultName = Left(defaultName, InStrRev(defaultName, ".") - 1)
    defaultName = defaultName & ".pdf"
    If pres.Path <> "" Then defaultName = pres.Path & "\" & defaultName

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Save as PDF"
        .InitialFileName = defaultName
        If .Show <> -1 Then Exit Sub ' User cancelled

        Dim pdfPath As String
        pdfPath = .SelectedItems(1)
    End With

    ' The dialog may add a PowerPoint extension (.pptx etc.) - strip it so the
    ' original deck can never be overwritten, then make sure the name ends in .pdf
    Dim dotPos As Long
    dotPos = InStrRev(pdfPath, ".")
    If dotPos > InStrRev(pdfPath, "\") Then
        If LCase(Mid(pdfPath, dotPos)) Like ".p[op]*" Then pdfPath = Left(pdfPath, dotPos - 1)
    End If
    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    pres.ExportAsFixedFormat pdfPath, ppFixedFormatTypePDF

    MsgBox "PDF saved to:" & vbNewLine & pdfPath, vbInformation
End Sub
