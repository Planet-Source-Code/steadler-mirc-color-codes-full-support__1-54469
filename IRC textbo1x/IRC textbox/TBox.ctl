VERSION 5.00
Begin VB.UserControl TBox 
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   DrawWidth       =   51
   KeyPreview      =   -1  'True
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   Begin VB.VScrollBar ScrollBar 
      Height          =   5520
      LargeChange     =   20
      Left            =   6720
      Max             =   1
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   1
      Width           =   315
   End
   Begin VB.PictureBox Display 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   5490
      Left            =   -15
      ScaleHeight     =   362
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   15
      Width           =   6675
   End
End
Attribute VB_Name = "TBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)

Dim ldispwidth As Long
'COnstants
Const ColorChr As String = ""
Const BoldChr As String = ""
Const PlainChr As String = ""
Const UnderlineChr As String = ""
Const ReverseChr As String = ""
Const MinAllowedWidth = 80
Const vbSpace As String = " "
Dim colorK(99) As Long
'Booleans
Dim Init As Boolean
Dim blnForceRef As Boolean

'Resize/draw variables
Dim lnSize As Integer
Dim maxLines As Integer

Dim rectPen As Long
Dim lngBrush As Long
Public sMax As Integer
Public sValue As Integer
Dim Words() As String
'Types
Private Type fStyle
    bold As Boolean
    underline As Boolean
    textColor As Long
    drect As Boolean
    rectColor As Long
End Type
Dim FontStyle As fStyle
Private Type dispParam
    belongs As Integer
    nLines As Integer
    TX As String
End Type
'Global ofelias
Dim SizeV As POINTAPI
Dim STable() As stringTable
Dim Fstrings(8000) As dispParam

Dim LineCount As Long
Dim LinesR As Integer
Dim startY As Integer
'VARIABLES FOR THE DISPLAYTEXT SUB
Dim lstBelong As Integer
  
'APIS VARIABLES
Dim Data() As Integer
Dim h(5) As Long
Dim StrProc As String
Dim backBuffer As Long
Dim backBuffer2 As Long
Dim lBitmap As Long
Dim lBitmap2 As Long
Dim pBackColor As Long
Dim lBackColor As Long
Dim lngJPG As Long
Dim winrect As RECT
'no in use

Public Sub NewLine(ByVal NewText As String)
    Dim tkLines As Integer
    Dim i As Integer
    Dim c As Integer
    'Add a new line
    LinesR = LinesR + 1
    
    'Find Color Codes
    If InStr(NewText, ColorChr) Then NewText = DefineColorChr(NewText)


    ReDim Preserve STable(LinesR)
    'Write to memory
    tkLines = SetTextMetrics(LinesR, NewText)

    
    For i = LineCount + 1 To LineCount + tkLines
        Fstrings(i).belongs = LinesR
        Fstrings(i).nLines = tkLines
    Next i
        LineCount = LineCount + tkLines
        blnForceRef = False
    
    If ScrollBar.Value = ScrollBar.Max Then
            ScrollBar.Max = LineCount
            ScrollBar.Value = LineCount
            sValue = ScrollBar.Value
    Else
            ScrollBar.Max = LineCount
    End If
        sMax = LineCount
    
        Init = True
        blnForceRef = True
     
    DisplayText
    
End Sub

Function SetTextMetrics(ByVal tLine As Integer, ByVal strText As String) As Integer
Dim i As Integer, iseek As Integer, res As Single, i2 As Integer
Dim parts() As String
STable(tLine).Words = strText
strText = Replace(strText, BoldChr, vbNullString)
strText = Replace(strText, PlainChr, vbNullString)
strText = Replace(strText, UnderlineChr, vbNullString)
strText = Replace(strText, ReverseChr, vbNullString)
If InStr(strText, ColorChr) > 0 Then

    Do

        iseek = InStr(iseek + 1, strText, ColorChr)
        If iseek > 0 Then
            strText = Mid(strText, 1, iseek - 1) & Mid(strText, iseek + 5, Len(strText) - iseek + 4)
        End If
    Loop Until iseek = 0
End If

        STable(tLine).wholeLen = GetTextWidth(strText)
        SetTextMetrics = Int(STable(tLine).wholeLen / Display.ScaleWidth) + IIf((STable(tLine).wholeLen Mod Display.ScaleWidth) = 0, 0, 1)

End Function

Public Function GetTextWidth(ByVal tText As String) As Integer
    GetTextExtentPoint32 Display.hdc, tText, Len(tText), SizeV
    GetTextWidth = SizeV.x
End Function


Private Sub Command1_Click()
NewLine "pasok6***Lookingupyourhostnament7disok1giolva"
End Sub

Private Sub Display_Resize()

Dim res As Single
Dim tkLines As Integer, i As Integer, i2 As Integer, Dispwidth As Long
Dispwidth = Display.ScaleWidth
If Dispwidth <> ldispwidth Then
    LineCount = 0
        For i = 1 To LinesR
            tkLines = Int(STable(i).wholeLen / Dispwidth) + IIf((STable(i).wholeLen Mod Dispwidth) = 0, 0, 1)
        For i2 = LineCount + 1 To LineCount + tkLines
    
            With Fstrings(i2)
                .belongs = i
                .nLines = tkLines
            End With
        Next i2
            LineCount = LineCount + tkLines
        Next i
 blnForceRef = False
If LineCount > 0 Then
    If ScrollBar.Value = ScrollBar.Max Then
   
        ScrollBar.Max = LineCount
        sValue = LineCount
        ScrollBar.Value = LineCount
    Else
   
        ScrollBar.Max = LineCount
    End If
    
    
End If
sMax = LineCount

End If

blnForceRef = True
ldispwidth = Dispwidth

DisplayText

End Sub
Public Function LoadTextSizes()
Dim ts As POINTAPI
Dim i As Long
    For i = 0 To 65525
        GetTextExtentPoint32 Display.hdc, ChrW(i), 1, ts
        Sizes(i) = ts.x
    Next i
    Display.FontBold = True
    For i = 0 To 65525
        GetTextExtentPoint32 Display.hdc, ChrW(i), 1, ts
        SizesB(i) = ts.x
    Next i
    
    maxLines = Int(Display.ScaleHeight / ts.y) + 1
    lnSize = ts.y
End Function

Private Sub ScrollBar_Change()
    If blnForceRef Then
        DisplayText
    End If
End Sub

Private Sub ScrollBar_Scroll()
  If blnForceRef Then
        DisplayText
    End If
End Sub

Private Sub UserControl_Initialize()
    
    pBackColor = RGB(255, 255, 255) 'For fucking dublicating shits and bullshit
    DeleteDC backBuffer
    DeleteObject lBitmap
   lBackColor = CreateSolidBrush(pBackColor)
    h(0) = 1
    h(1) = 2
    h(3) = StrPtr(StrProc)
    h(4) = &H7FFFFFFF
    RtlMoveMemory ByVal ArrPtr(Data), VarPtr(h(0)), 4
  '  backBuffer2 = CreateCompatibleDC(GetDC(0))
'SelectObject backBuffer2, lBackColor
 

 '   backBuffer = CreateCompatibleDC(GetDC(0))
  '      lBitmap2 = CreateCompatibleBitmap(GetDC(0), Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)

    'lBitmap = CreateCompatibleBitmap(GetDC(0), Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
   'SelectObject backBuffer, lBitmap
    '  SelectObject backBuffer2, lBitmap

   
   ' lBackColor = CreateSolidBrush(pBackColor)
    With winrect
    .Left = 0
    .Top = 0
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
    .Right = Screen.Width / Screen.TwipsPerPixelX
    End With
  '  FillRect backBuffer2, winrect, lBackColor
   ' SetBkMode backBuffer, 1
colorK(0) = RGB(255, 255, 255)
colorK(1) = RGB(0, 0, 0)
colorK(2) = RGB(0, 0, 127)
colorK(3) = RGB(0, 147, 0)
colorK(4) = RGB(255, 0, 0)
colorK(5) = RGB(127, 0, 0)
colorK(6) = RGB(156, 0, 156)
colorK(7) = RGB(252, 127, 0)
colorK(8) = RGB(255, 255, 0)
colorK(9) = RGB(0, 252, 0)
colorK(10) = RGB(0, 147, 147)
colorK(11) = RGB(0, 255, 255)
colorK(12) = RGB(0, 0, 252)
colorK(13) = RGB(255, 0, 255)
colorK(14) = RGB(127, 127, 127)
colorK(15) = RGB(210, 210, 210)
colorK(99) = RGB(255, 255, 255)
    LoadTextSizes
End Sub

Private Sub UserControl_Resize()
With Display
    .Width = UserControl.ScaleWidth - ScrollBar.Width
    .Height = UserControl.ScaleHeight
End With
With ScrollBar
.Height = UserControl.ScaleHeight
.Left = Display.Width
End With
maxLines = Int(Display.ScaleHeight / lnSize) + 1

End Sub

Public Sub DisplayText()

Dim Lines2Draw As Integer, i As Integer, i2 As Integer, dx As Integer, dx2 As Integer, li As Integer, lbelong As Integer, lp As Integer, isB As Boolean
    If Init And blnForceRef Then
    winrect.Bottom = Display.ScaleHeight
    winrect.Right = Display.ScaleWidth
    FillRect Display.hdc, winrect, lBackColor
    'Clear BOX ;Poses grammes?
    For i = ScrollBar.Value To ScrollBar.Value - maxLines Step -1
        If Sgn(i) = 1 Then
            Lines2Draw = Lines2Draw + 1
        End If
    Next i
    startY = Display.ScaleHeight - lnSize * Lines2Draw
    formatS ScrollBar.Value - Lines2Draw + 1, ScrollBar.Value, Display.ScaleWidth, Display.hdc
    
    For i = ScrollBar.Value - Lines2Draw + 1 To ScrollBar.Value
        'Set standar setting
        dx = 0
        dx2 = 0
        li = 0
        lp = -1
        If lbelong <> Fstrings(i).belongs Then
        With FontStyle
            .textColor = colorK(1)
            .drect = False
            .bold = False
            .underline = False
        End With
        isB = False
         Display.FontBold = False
        Display.FontUnderline = False
        SetBkColor Display.hdc, colorK(0)
        SetTextColor Display.hdc, colorK(1)
        lbelong = Fstrings(i).belongs
        End If
        If Fstrings(i).TX <> vbNullString Then
        
         StrProc = Fstrings(i).TX
         h(3) = StrPtr(StrProc)
        For i2 = 0 To Len(StrProc)
            If Data(i2) = 3 Then
                TextOut Display.hdc, dx2, startY, Mid$(Fstrings(i).TX, li + 1, i2 - li), Len(Mid$(Fstrings(i).TX, li + 1, i2 - li))
                If Not Val(Mid$(Fstrings(i).TX, i2 + 4, 2)) = 99 Then
                SetBkColor Display.hdc, colorK(Val(Mid$(Fstrings(i).TX, i2 + 4, 2)))
                End If
                If Val(Mid$(Fstrings(i).TX, i2 + 2, 2)) = 99 Then
                    SetBkColor Display.hdc, colorK(1)
                Else
                    SetTextColor Display.hdc, colorK(Val(Mid$(Fstrings(i).TX, i2 + 2, 2)))
                End If
                dx2 = dx
                i2 = i2 + 4
                lp = i2
                li = i2 + 1
            ElseIf Data(i2) = 2 Then
                TextOut Display.hdc, dx2, startY, Mid$(Fstrings(i).TX, li + 1, i2 - li), Len(Mid$(Fstrings(i).TX, li + 1, i2 - li))
                Display.FontBold = Not Display.FontBold
                isB = Not isB
                lp = i2
                li = i2 + 1
                dx2 = dx
                
            ElseIf Data(i2) = 31 Then
                TextOut Display.hdc, dx2, startY, Mid$(Fstrings(i).TX, li + 1, i2 - li), Len(Mid$(Fstrings(i).TX, li + 1, i2 - li))
                Display.FontUnderline = Not Display.FontUnderline
                lp = i2
                li = i2 + 1
                dx2 = dx
            Else
                
                dx = dx + IIf(isB, SizesB(i2), Sizes(i2))
            End If
        Next i2
         End If
        TextOut Display.hdc, dx2, startY, Mid$(Fstrings(i).TX, lp + 2, i2 - lp), Len(Mid$(Fstrings(i).TX, lp + 2, i2 - lp))
         startY = startY + lnSize
    Next i
    Display.Refresh
    End If

End Sub
Public Sub formatS(ByVal Ei As Integer, ByVal Si As Integer, ByVal Dispwidth As Integer, ByVal hdc As Long)
'fix ei FIX CODE

Dim i As Integer, hbelong As Integer, iW As Integer, wLen As Integer, totalLen As Integer, dl As Integer, liW As Integer, wWidth As Integer, remain As Integer, wWS As Integer, wClosed As Boolean, isB As Boolean
    hbelong = Fstrings(Ei).belongs
For i = Ei To 0 Step -1
    If hbelong <> Fstrings(Ei).belongs Then
        Ei = i + 1
        i = -1
    End If
Next i
    hbelong = 0
For i = Ei To Si
    If hbelong <> Fstrings(i).belongs Then
        isB = False
        'AN EINE OK TOTE STEILTO
        If STable(Fstrings(i).belongs).wholeLen < Dispwidth Then
            Fstrings(i).TX = STable(Fstrings(i).belongs).Words
        Else
            StrProc = STable(Fstrings(i).belongs).Words
            h(3) = StrPtr(StrProc)
            dl = 0
            liW = 1
            wWidth = 0
            wWS = 0
            totalLen = 0
            remain = Fstrings(i).nLines * Dispwidth - STable(Fstrings(i).belongs).wholeLen
           
            For iW = 0 To Len(StrProc) - 1
                If Data(iW) = 3 Then
                     iW = iW + 4
                     wWS = wWS + 4
                ElseIf Data(iW) = 2 Then
                isB = Not isB
                ElseIf Data(iW) = 15 Or Data(iW) = 31 Or Data(iW) = 22 Then
                    'nothing
                Else
                    If Data(iW) = 32 Then wWidth = Sizes(32): wWS = 0
                    wWS = wWS + 1
                    On Error Resume Next
                    totalLen = totalLen + Sizes(Data(iW))
                    wWidth = wWidth + Sizes(Data(iW))
                    If totalLen > Dispwidth And dl < Fstrings(i).nLines Then
                      If remain - wWidth >= 10000 Then
                        Fstrings(i + dl).TX = Mid$(StrProc, liW, iW - liW - wWS + 3)
                        remain = remain - wWidth
                        dl = dl + 1
                        liW = iW + 2 - wWS
                        iW = iW - wWS + 2
                        wWS = 0
                        wWidth = 0
                        totalLen = Sizes(32)
                        liW = iW + 1
                      Else
                        Fstrings(i + dl).TX = Mid$(StrProc, liW, iW - liW + 1)
                        wWS = 0
                        wWidth = 0
                        liW = iW + 1
                        totalLen = Sizes(Data(iW))
                        
                        dl = dl + 1
                      End If
                    End If
                End If
                
                
                
            Next iW
           
            Fstrings(i + dl).TX = Mid$(StrProc, liW, iW - liW + 1)
            
        End If
        hbelong = Fstrings(i).belongs
    End If
Next i
Exit Sub
errs:
MsgBox Err.Description
End Sub




Private Sub UserControl_Terminate()

RtlMoveMemory ByVal ArrPtr(Data), 0&, 4
    DeleteObject rectPen
    DeleteObject FontStyle.rectColor
    DeleteDC backBuffer
    DeleteObject lBitmap
    DeleteObject lBackColor
End Sub

Public Sub UpdateBar()
blnForceRef = True
ScrollBar.Value = sValue

blnForceRef = False
End Sub
