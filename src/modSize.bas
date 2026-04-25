Attribute VB_Name = "modSize"
' ============================================================
' modSize — Size and stretch commands
' ============================================================

' --- Match to first selected ---

Public Sub MatchWidth(control As IRibbonControl)
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refW As Single
    refW = sr(1).Width
    Dim i As Integer

    For i = 2 To sr.Count
        sr(i).Width = refW
    Next i
End Sub

Public Sub MatchHeight(control As IRibbonControl)
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refH As Single
    refH = sr(1).Height
    Dim i As Integer

    For i = 2 To sr.Count
        sr(i).Height = refH
    Next i
End Sub

Public Sub MatchWidthAndHeight(control As IRibbonControl)
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refW As Single, refH As Single
    refW = sr(1).Width
    refH = sr(1).Height
    Dim i As Integer

    For i = 2 To sr.Count
        sr(i).Width = refW
        sr(i).Height = refH
    Next i
End Sub

' --- Copy / Paste size ---

Public Sub CopySize(control As IRibbonControl)
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one shape.", vbExclamation
        Exit Sub
    End If

    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Please select one shape.", vbExclamation
        Exit Sub
    End If

    storedWidth = ActiveWindow.Selection.ShapeRange(1).Width
    storedHeight = ActiveWindow.Selection.ShapeRange(1).Height
    hasCopiedSize = True
End Sub

Public Sub PasteWidth(control As IRibbonControl)
    If Not hasCopiedSize Then Exit Sub

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim i As Integer

    For i = 1 To sr.Count
        sr(i).Width = storedWidth
    Next i
End Sub

Public Sub PasteHeight(control As IRibbonControl)
    If Not hasCopiedSize Then Exit Sub

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim i As Integer

    For i = 1 To sr.Count
        sr(i).Height = storedHeight
    Next i
End Sub

Public Sub PasteWidthAndHeight(control As IRibbonControl)
    If Not hasCopiedSize Then Exit Sub

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    Dim i As Integer

    For i = 1 To sr.Count
        sr(i).Width = storedWidth
        sr(i).Height = storedHeight
    Next i
End Sub

' --- Stretch commands ---
' In each stretch command:
'   sr(1) is the reference shape
'   All other shapes have one edge pulled to meet the reference boundary
'   The opposite edge stays fixed (achieved by adjusting Top/Left AND Height/Width together)

Public Sub StretchUp(control As IRibbonControl)
    ' Pulls the top edge of each shape up to meet the bottom edge of the reference shape
    ' Bottom edge of each shape stays fixed
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refBottom As Single
    refBottom = sr(1).Top + sr(1).Height
    Dim i As Integer

    For i = 2 To sr.Count
        Dim oldBottom As Single
        oldBottom = sr(i).Top + sr(i).Height
        sr(i).Top = refBottom
        sr(i).Height = oldBottom - refBottom
    Next i
End Sub

Public Sub StretchDown(control As IRibbonControl)
    ' Pulls the bottom edge of each shape down to meet the top edge of the reference shape
    ' Top edge of each shape stays fixed
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refTop As Single
    refTop = sr(1).Top
    Dim i As Integer

    For i = 2 To sr.Count
        sr(i).Height = refTop - sr(i).Top
    Next i
End Sub

Public Sub StretchLeft(control As IRibbonControl)
    ' Pulls the left edge of each shape to meet the right edge of the reference shape
    ' Right edge of each shape stays fixed
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refRight As Single
    refRight = sr(1).Left + sr(1).Width
    Dim i As Integer

    For i = 2 To sr.Count
        Dim oldRight As Single
        oldRight = sr(i).Left + sr(i).Width
        sr(i).Left = refRight
        sr(i).Width = oldRight - refRight
    Next i
End Sub

Public Sub StretchRight(control As IRibbonControl)
    ' Pulls the right edge of each shape to meet the left edge of the reference shape
    ' Left edge of each shape stays fixed
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange

    If sr.Count < 2 Then
        MsgBox "Please select two or more shapes.", vbExclamation
        Exit Sub
    End If

    Dim refLeft As Single
    refLeft = sr(1).Left
    Dim i As Integer

    For i = 2 To sr.Count
        sr(i).Width = refLeft - sr(i).Left
    Next i
End Sub
