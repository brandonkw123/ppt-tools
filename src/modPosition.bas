Attribute VB_Name = "modPosition"
' ============================================================
' modPosition - Position-related commands
' ============================================================

Public Sub SwapPositions(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(2, 2)
    If sr Is Nothing Then Exit Sub

    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single

    x1 = sr(1).Left
    y1 = sr(1).Top
    x2 = sr(2).Left
    y2 = sr(2).Top

    sr(1).Left = x2
    sr(1).Top = y2
    sr(2).Left = x1
    sr(2).Top = y1
End Sub

Public Sub CopyPosition(control As IRibbonControl)
    Dim sr As ShapeRange
    Set sr = SelectedShapes(1, 1)
    If sr Is Nothing Then Exit Sub

    storedLeft = sr(1).Left
    storedTop = sr(1).Top
    hasCopiedPosition = True
End Sub

Public Sub PastePosition(control As IRibbonControl)
    If Not hasCopiedPosition Then
        Beep
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = SelectedShapes(1)
    If sr Is Nothing Then Exit Sub

    Dim i As Integer
    For i = 1 To sr.Count
        sr(i).Left = storedLeft
        sr(i).Top = storedTop
    Next i
End Sub
