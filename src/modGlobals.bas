Attribute VB_Name = "modGlobals"
' ============================================================
' modGlobals - Shared variables and helpers across all modules
' ============================================================

' Stored position values
Public storedLeft As Single
Public storedTop As Single

' Stored size values
Public storedWidth As Single
Public storedHeight As Single

' Flags to track whether a copy has been performed
Public hasCopiedPosition As Boolean
Public hasCopiedSize As Boolean

' Returns the selected shapes if there are between minCount and maxCount of them.
' Otherwise plays the Windows error sound and returns Nothing.
' A cursor inside a text box counts as selecting that shape.
Public Function SelectedShapes(minCount As Integer, Optional maxCount As Integer = 9999) As ShapeRange
    Dim valid As Boolean

    If Application.Windows.Count > 0 Then
        Select Case ActiveWindow.Selection.Type
            Case ppSelectionShapes, ppSelectionText
                With ActiveWindow.Selection.ShapeRange
                    valid = (.Count >= minCount And .Count <= maxCount)
                End With
        End Select
    End If

    If valid Then
        Set SelectedShapes = ActiveWindow.Selection.ShapeRange
    Else
        Beep
    End If
End Function
