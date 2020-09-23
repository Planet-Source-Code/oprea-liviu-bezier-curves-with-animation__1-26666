Attribute VB_Name = "Module1"
Public Declare Function PolyBezier Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long


Public Type POINTAPI
        x As Long
        y As Long
End Type

Public c As Integer
