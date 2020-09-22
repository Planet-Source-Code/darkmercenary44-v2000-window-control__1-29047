Attribute VB_Name = "Module1"
Option Explicit

Public ism As Boolean

Public X1 As Long
Public X2 As Long
Public Y1 As Long
Public Y2 As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public IsInRegion As Boolean
Public Pos As POINTAPI
Public MMove As POINTAPI
Public SMove As POINTAPI
Public TMove As POINTAPI

Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


