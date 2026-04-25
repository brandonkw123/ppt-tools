Attribute VB_Name = "modGlobals"
' ============================================================
' modGlobals — Shared variables across all modules
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
