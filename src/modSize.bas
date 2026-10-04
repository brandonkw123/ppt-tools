Attribute VB_Name = "modSize"
' ============================================================
' modSize - Size and stretch commands
' ============================================================

' Sets a shape's width and/or height (pass -1 to leave one unchanged).
' Lock Aspect Ratio is switched off while resizing so only the requested
' dimension changes, then restored.
Private Sub SetSize(shp As Shape, ByVal newW As Single, ByVal newH As Single)
    Dim lockState As MsoTriState
    lockState = shp.LockAspectRatio
    shp.LockAspectRatio = msoFalse
    If newW >= 0 Then shp.Width = newW
    If newH >= 0 Then shp.Height = newH
    shp.LockAspectRatio = lockState
End Sub

' --- Match to first selected ---

Public Sub MatchWidth(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(2)
    If sr Is Nothing Then Exit Sub

    Dim i As Integer
    For i = 2 To sr.Count
        SetSize sr(i), sr(1).Width, -1
    Next i
End Sub

Public Sub MatchHeight(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(2)
    If sr Is Nothing Then Exit Sub

    Dim i As Integer
    For i = 2 To sr.Count
        SetSize sr(i), -1, sr(1).Height
    Next i
End Sub

Public Sub MatchWidthAndHeight(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(2)
    If sr Is Nothing Then Exit Sub

    Dim i As Integer
    For i = 2 To sr.Count
        SetSize sr(i), sr(1).Width, sr(1).Height
    Next i
End Sub

' --- Copy / Paste size ---

Public Sub CopySize(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(1, 1)
    If sr Is Nothing Then Exit Sub

    storedWidth = sr(1).Width
    storedHeight = sr(1).Height
    hasCopiedSize = True
End Sub

Public Sub PasteWidth(control As IRibbonControl)
    PasteSize True, False
End Sub

Public Sub PasteHeight(control As IRibbonControl)
    PasteSize False, True
End Sub

Public Sub PasteWidthAndHeight(control As IRibbonControl)
    PasteSize True, True
End Sub

Private Sub PasteSize(doWidth As Boolean, doHeight As Boolean)
    If Not hasCopiedSize Then
        Beep
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = SelectedShapes(1)
    If sr Is Nothing Then Exit Sub

    Dim i As Integer
    For i = 1 To sr.Count
        SetSize sr(i), IIf(doWidth, storedWidth, -1), IIf(doHeight, storedHeight, -1)
    Next i
End Sub

' --- Stretch commands ---
' sr(1) - the selected shape furthest back - is the reference shape.
' Every other selected shape has one edge moved to a line on the reference;
' the opposite edge stays fixed.
'   "to Meet":  the edge moves to the reference's NEAR edge, so the shapes touch
'               (Stretch Right to Meet: right edge -> reference's left edge)
'   "to Match": the edge moves to the reference's SAME edge, so the edges line up
'               (Stretch Right to Match: right edge -> reference's right edge)
' If that line is on the wrong side of a shape's fixed edge, the shape would need
' zero or negative size, so it is skipped and the error sound plays.

Public Sub StretchUp(control As IRibbonControl)
    StretchShapes "Up", False
End Sub

Public Sub StretchDown(control As IRibbonControl)
    StretchShapes "Down", False
End Sub

Public Sub StretchLeft(control As IRibbonControl)
    StretchShapes "Left", False
End Sub

Public Sub StretchRight(control As IRibbonControl)
    StretchShapes "Right", False
End Sub

Public Sub StretchUpMatch(control As IRibbonControl)
    StretchShapes "Up", True
End Sub

Public Sub StretchDownMatch(control As IRibbonControl)
    StretchShapes "Down", True
End Sub

Public Sub StretchLeftMatch(control As IRibbonControl)
    StretchShapes "Left", True
End Sub

Public Sub StretchRightMatch(control As IRibbonControl)
    StretchShapes "Right", True
End Sub

Private Sub StretchShapes(direction As String, toMatch As Boolean)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(2)
    If sr Is Nothing Then Exit Sub

    ' The line on the reference shape that the moving edge goes to
    Dim ref As Shape, target As Single
    Set ref = sr(1)
    Select Case direction
        Case "Up":    target = IIf(toMatch, ref.Top, ref.Top + ref.Height)
        Case "Down":  target = IIf(toMatch, ref.Top + ref.Height, ref.Top)
        Case "Left":  target = IIf(toMatch, ref.Left, ref.Left + ref.Width)
        Case "Right": target = IIf(toMatch, ref.Left + ref.Width, ref.Left)
    End Select

    Dim i As Integer, shp As Shape, newSize As Single, skipped As Boolean

    For i = 2 To sr.Count
        Set shp = sr(i)
        Select Case direction
            Case "Up"       ' bottom edge stays fixed
                newSize = shp.Top + shp.Height - target
                If newSize > 0 Then
                    shp.Top = target
                    SetSize shp, -1, newSize
                Else
                    skipped = True
                End If
            Case "Down"     ' top edge stays fixed
                newSize = target - shp.Top
                If newSize > 0 Then SetSize shp, -1, newSize Else skipped = True
            Case "Left"     ' right edge stays fixed
                newSize = shp.Left + shp.Width - target
                If newSize > 0 Then
                    shp.Left = target
                    SetSize shp, newSize, -1
                Else
                    skipped = True
                End If
            Case "Right"    ' left edge stays fixed
                newSize = target - shp.Left
                If newSize > 0 Then SetSize shp, newSize, -1 Else skipped = True
        End Select
    Next i

    If skipped Then Beep
End Sub
