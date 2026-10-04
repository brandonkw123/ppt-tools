Attribute VB_Name = "modCleanup"
' ============================================================
' modCleanup - Remove comments and speaker notes from the whole deck
' ============================================================

Public Sub RemoveAllComments(control As IRibbonControl)
    RunCleanup True, False
End Sub

Public Sub RemoveAllNotes(control As IRibbonControl)
    RunCleanup False, True
End Sub

Public Sub RemoveCommentsAndNotes(control As IRibbonControl)
    RunCleanup True, True
End Sub

' Saves a backup copy of the deck, then removes comments and/or notes from every slide
Private Sub RunCleanup(doComments As Boolean, doNotes As Boolean)
    If Presentations.Count = 0 Then
        Beep
        Exit Sub
    End If

    Dim pres As Presentation
    Set pres = ActivePresentation

    Dim sld As Slide
    Dim found As Boolean

    For Each sld In pres.Slides
        If doComments And sld.Comments.Count > 0 Then found = True
        If doNotes Then
            If NotesText(sld) <> "" Then found = True
        End If
    Next sld

    If Not found Then
        MsgBox "Nothing to remove - no " & WhatText(doComments, doNotes) & " found.", vbInformation
        Exit Sub
    End If

    ' Save a backup copy before changing anything; stop if that fails
    Dim backupPath As String
    backupPath = SaveBackup(pres)
    If backupPath = "" Then
        MsgBox "Couldn't save a backup copy, so nothing was removed.", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    Dim body As Shape

    For Each sld In pres.Slides
        If doComments Then
            ' Delete backwards so the remaining indexes don't shift
            For i = sld.Comments.Count To 1 Step -1
                sld.Comments(i).Delete
            Next i
        End If

        If doNotes Then
            Set body = NotesBody(sld)
            If Not body Is Nothing Then body.TextFrame.TextRange.Text = ""
        End If
    Next sld

    MsgBox "Successfully removed. Backup copy saved to temp folder.", vbInformation
End Sub

' Saves a copy of the deck to %TEMP%\PPT Tools backups and returns its path ("" if it failed)
Private Function SaveBackup(pres As Presentation) As String
    Dim folder As String
    folder = Environ("TEMP") & "\PPT Tools backups"
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Dim baseName As String
    baseName = pres.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)

    Dim backupFile As String
    backupFile = folder & "\" & baseName & " (before cleanup " & Format(Now, "yyyy-mm-dd hh.nn.ss") & ").pptx"

    On Error Resume Next
    pres.SaveCopyAs backupFile, ppSaveAsOpenXMLPresentation
    If Err.Number = 0 Then SaveBackup = backupFile
    On Error GoTo 0
End Function

' Returns the speaker notes text box on a slide's notes page, or Nothing
Private Function NotesBody(sld As Slide) As Shape
    Dim shp As Shape
    For Each shp In sld.NotesPage.Shapes.Placeholders
        If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
            If shp.HasTextFrame Then
                Set NotesBody = shp
                Exit Function
            End If
        End If
    Next shp
End Function

Private Function NotesText(sld As Slide) As String
    Dim body As Shape
    Set body = NotesBody(sld)
    If Not body Is Nothing Then NotesText = Trim(body.TextFrame.TextRange.Text)
End Function

Private Function WhatText(doComments As Boolean, doNotes As Boolean) As String
    If doComments And doNotes Then
        WhatText = "comments or speaker notes"
    ElseIf doComments Then
        WhatText = "comments"
    Else
        WhatText = "speaker notes"
    End If
End Function
