Attribute VB_Name = "ColorFunction"
Option Explicit
Dim parts() As String
Public Const ColorChr As String = ""
Public Const BoldChr As String = ""
Public Const PlainChr As String = ""
Public Const UnderlineChr As String = ""
Public Const ReverseChr As String = ""
Public Const vbSpace As String = " "


Public Function DefineColorChr(ByVal sText As String) As String

Dim i As Integer, i2 As Integer, returnS As String, t As String, t2 As String, l As Byte
parts = Split(sText, ColorChr)
returnS = parts(0)
    For i = 1 To UBound(parts)
        Select Case True
            Case parts(i) Like "##,##*" Or parts(i) Like "##,#*"
                 If parts(i) Like "##,##*" Then l = 2 Else l = 1
                 t = Mid(parts(i), 1, 2)
                 t2 = Mid(parts(i), 4, l)
                 parts(i) = Replace(parts(i), t & "," & t2, vbNullString, , 1)
                 returnS = returnS & ColorChr & LZ(t) & LZ(t2) & parts(i)
            Case parts(i) Like "#,##*" Or parts(i) Like "#,#*"
                 If parts(i) Like "#,##*" Then l = 2 Else l = 1
                 t = Mid(parts(i), 1, 1)
                 t2 = Mid(parts(i), 3, l)
                 parts(i) = Replace(parts(i), t & "," & t2, vbNullString, , 1)
                 returnS = returnS & ColorChr & LZ(t) & LZ(t2) & parts(i)
                 
            Case parts(i) Like "#*" Or parts(i) Like "##*"
                 If parts(i) Like "##*" Then l = 2 Else l = 1
                 t = Mid(parts(i), 1, l)
                 parts(i) = Replace(parts(i), t, vbNullString, , 1)
                 returnS = returnS & ColorChr & LZ(t) & "99" & parts(i)
            Case Else
            
            returnS = returnS & ColorChr & "99" & "99" & parts(i)
        
        End Select
    Next
    DefineColorChr = returnS
End Function
Function LZ(ByVal St As String) As String
   LZ = String(2 - Len(St), "0") & St
End Function
