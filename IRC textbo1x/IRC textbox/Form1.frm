VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   8040
      TabIndex        =   3
      Top             =   0
      Width           =   15
   End
   Begin vb6projectProject1.TBox TBox1 
      Height          =   6105
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      _extentx        =   14208
      _extenty        =   10769
   End
   Begin VB.ListBox nlist 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      IntegralHeight  =   0   'False
      Left            =   8085
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   15
      Width           =   1260
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   7980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mQuotes(20) As String
Dim nQ As Byte
Dim nQ2 As Byte
Dim numofQ As Byte



Private Sub Command1_Click()
TBox1.NewLine "*** Looking up your hostname *** Looking up your hostname *** Looking up your hostname"

End Sub

Private Sub Form_Initialize()
Form_Resize
End Sub



Private Sub Form_Resize()
On Local Error Resume Next
    TBox1.Width = Me.ScaleWidth - nlist.Width
    nlist.Height = Me.ScaleHeight - Text1.Height
    nlist.Left = TBox1.Width
   TBox1.Height = Me.ScaleHeight - Text1.Height
   Frame1.Left = TBox1.Width
   Frame1.Height = TBox1.Height
   Text1.Top = TBox1.Height
   Text1.Width = Me.ScaleWidth
End Sub

Private Sub luser_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
MDIForm1.socket.SendData "PART " & Me.Tag & vbCrLf
Do Until formes(y) = Me.Tag
If formes(y) = Me.Tag Then formes(y) = ""
y = y + 1
Loop
Me.Tag = ""

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Frame1.MousePointer = 9
If Button = 1 Then
On Error Resume Next
    Frame1.Left = Frame1.Left + x / Screen.TwipsPerPixelX
    TBox1.Width = Frame1.Left
    nlist.Left = Frame1.Left
    nlist.Width = Me.ScaleWidth - nlist.Left
    DoEvents
End If
End Sub

Private Sub TBox1_GotFocus()
'Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   If Text1.Text = "/l" Then
   MDIForm1.Nick = "XNeon2"
     MDIForm1.socket.SendData "NICK XNeon2" & vbCrLf
      MDIForm1.socket.SendData "USER dddd dddd dddd :dddd" & vbCrLf
   End If
   If Text1.Text = "/h" Then
      HookWindowEvents Me.hwnd
   End If
   If Text1.Text = "/u" Then
     UnhookWindowEvents
   End If
   If Text1.Text = "/lo" Then
     MDIForm1.socket.SendData "PRIVMSG X@channels.undernet.org :login igv irbmesucks" & vbCrLf
        MDIForm1.socket.SendData "MODE " & MDIForm1.Nick & " +x" & vbCrLf

   End If
If Left(Text1.Text, 5) = "/join" Then

MDIForm1.socket.SendData "join " & Split(Text1.Text)(1) & vbCrLf
End If
If Left(Text1.Text, 5) = "/nick" Then

MDIForm1.socket.SendData "NICK " & Split(Text1.Text)(1) & vbCrLf
End If
If Left(Text1.Text, 5) = "/part" Then

MDIForm1.socket.SendData "PART " & Split(Text1.Text)(1) & vbCrLf
End If
   mQuotes(nQ2) = Text1.Text
   nQ2 = nQ2 + 1
   numofQ = numofQ + 1
   If nQ2 >= 21 Then
      nQ2 = 0
   End If
   If numofQ > 20 Then numofQ = 20
   nQ = nQ2
   If Not Left(Text1.Text, 1) = "/" Then
   MDIForm1.socket.SendData "PRIVMSG " & Me.Tag & " :" & Text1.Text & vbCrLf
    TBox1.NewLine "<" & MDIForm1.Nick & "> " & Text1.Text
    End If
    '  mdiform1.socket.SendData Text1.Text & vbCrLf

  
   Text1.Text = ""
End If

If KeyCode = vbKeyUp Then
   If nQ > 0 Then
      nQ = nQ - 1
   Else
      nQ = 20
   End If
   Text1.Text = mQuotes(nQ)
End If
If KeyCode = vbKeyDown Then
   If nQ < numofQ Then
      nQ = nQ + 1
   Else
      nQ = numofQ
   End If
   Text1.Text = mQuotes(nQ)
End If
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
TBox1.NewLine Text2.Text
Text2.Text = ""
End If

End Sub

