Attribute VB_Name = "modPosition"
' ============================================================
' modPosition — Position-related commands
' ============================================================

Public Sub SwapPositions(control As IRibbonControl)
    Dim sr As ShapeRange
    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select exactly two shapes.", vbExclamation
        Exit Sub
    End If

    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count <> 2 Then
        MsgBox "Please select exactly two shapes.", vbExclamation
        Exit Sub
    End If

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
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one shape.", vbExclamation
        Exit Sub
    End If

    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Please select one shape.", vbExclamation
        Exit Sub
    End If

    storedLeft = ActiveWindow.Selection.ShapeRange(1).Left
    storedTop = ActiveWindow.Selection.ShapeRange(1).Top
    hasCopiedPosition = True
End Sub

Public Sub PastePosition(control As IRibbonControl)
    If Not hasCopiedPosition Then Exit Sub

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim i As Integer

    For i = 1 To sr.Count
        sr(i).Left = storedLeft
        sr(i).Top = storedTop
    Next i
End Sub
