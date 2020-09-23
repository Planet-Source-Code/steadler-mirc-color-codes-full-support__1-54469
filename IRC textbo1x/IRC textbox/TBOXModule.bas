Attribute VB_Name = "TBOXModule"
Option Explicit
Const vbSpace = " "
Private Type dispParam
belongs As Integer
nLines As Integer
iLine As Byte
TX As String
End Type
Dim SizeV As POINTAPI

Public Type tBuffer
    sString As String
    mBelongs As Integer
End Type
Public Type stringTable
    Words As String
    wholeLen As Integer
End Type


Const ColorChr As String = ""
Const BoldChr As String = ""
Const PlainChr As String = ""
Const UnderlineChr As String = ""
Const ReverseChr As String = ""
Const MinAllowedWidth = 80

Dim SpW() As String

Public Function GetTextWidth2(ByVal tText As String, ByVal hw As Long) As Integer
GetTextExtentPoint32 hw, tText, Len(tText), SizeV
GetTextWidth2 = SizeV.x
End Function

