VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormZC() As Form1

Dim i2 As Integer, stringData As String
Public Nick As String

Private Sub MDIForm_Load()

ReDim Preserve FormZC(i2)
Set FormZC(i2) = New Form1
FormZC(i2).Show
FormZC(i2).Tag = "Server"
formes(i2) = "Server"
i2 = i2 + 1

socket.Connect "Princeton.NJ.US.Undernet.Org", 6667
End Sub
Private Sub socket_DataArrival(ByVal bytesTotal As Long)
Dim sp() As String
Dim l() As String
socket.GetData stringData


l = Split(stringData, vbCrLf)
For i = 0 To UBound(l) - 1

sp = Split(l(i), " ")

Select Case sp(0)

Case "NOTICE"
   Drawtext "-5" & Split(l(i), ":")(1), "Server"

End Select
On Error Resume Next
Select Case Val(sp(1))

Case 353

ch = sp(4)
TX = Split(l(i), ":", 3)(2)
nic = Split(TX, " ")

For S = 0 To UBound(nic) - 1

Adduser nic(S), ch

Next S
Case 366

Case 332 'TOPIC

Drawtext "3* Topic is '" & Split(l(i), ":", 3)(2) & "'", sp(3)

Case 333

Case 1 To 999

sp(0) = vbNullString: sp(1) = vbNullString: sp(2) = vbNullString
Drawtext Replace(Replace(Join(sp, " "), "   ", "", 1, 1), ":", vbNullString), "Server"

End Select

Select Case sp(1)


Case "NOTICE"
Drawtext "-", "Server"
Drawtext "5-" & sp(0) & ": " & Split(l(i), ":", 3)(2), "Server"
Drawtext "-", "Server"

Case "JOIN"
If Replace(Split(l(i), "!")(0), ":", "") = Nick Then
MDIForm1.NewFRM sp(2)
Drawtext "12* Now talking in " & sp(2), sp(2)
Else
Drawtext "3* " & Replace(Split(l(i), "!")(0), ":", "") & " has joined " & sp(2), sp(2)
Adduser Replace(Split(l(i), "!")(0), ":", ""), sp(2)
End If

Case "PART"
Drawtext "3* " & Replace(Split(l(i), "!")(0), ":", "") & " has left " & sp(2), sp(2)
Remuser Replace(Split(l(i), "!")(0), ":", ""), sp(2)

Case "MODE"

Drawtext "3* " & Replace(Split(l(i), "!")(0), ":", "") & " sets mode: " & Split(l(i), " ", 4)(3), sp(2)


Case "QUIT"

'Drawtext "2* " & Replace(Split(l(I), "!")(0), ":", "") & " has quited " & sp(2), sp(1)

Case "NICK"

tni = Replace(Split(l(i), "!")(0), ":", "")

If tni = Nick Then

Nick = Replace(sp(2), ":", "")

End If
Drawtext "7* " & tni & " changed nick to " & Replace(sp(2), ":", ""), "*", tni

Case "PRIVMSG"
On Error Resume Next
msg = Split(l(i), ":", 3)(2)
ni = Replace(Split(l(i), "!")(0), ":", "")
If InStr(msg, "ACTION") Then
Drawtext "6* " & ni & "" & Replace(Replace(msg, "ACTION", ""), "", ""), sp(2)



Else


Drawtext "<" & ni & "> " & msg, sp(2)


End If


End Select


If Split(l(i))(0) = "PING" Then socket.SendData "PONG " & Split(stringData, " ")(1) & vbCrLf: Drawtext "-", "Server": Drawtext ByVal "-3Ping?Pong!", "Server": Drawtext "-", "Server"

Next i
End Sub

Public Function NewFRM(ByVal channel As String)
ReDim Preserve FormZC(i2)
Set FormZC(i2) = New Form1
FormZC(i2).Tag = channel
formes(i2) = channel

i2 = i2 + 1
End Function

Public Function Drawtext(ByVal tex As String, ByVal forma As String, Optional ByVal nisk As String)


If forma = Nick Then

ElseIf forma = "*" Then
   
    Do Until formes(x) = vbNullString
  
    For g = 0 To FormZC(x).nlist.ListCount
   
    If FormZC(x).nlist.List(g) = nisk Or formes(x) = "Server" Then
   
    FormZC(x).TBox1.NewLine tex
    Dim o
    For o = 0 To FormZC(x).nlist.ListCount - 1
    
    If FormZC(x).nlist.List(o) = nisk Then
    FormZC(x).nlist.List(o) = Split(tex, " ")(5)
    End If
    Next o
    
    
    End If
    Next g
    
x = x + 1
Loop


Else




Do Until p
If formes(x) = forma Then

If formes(x) = FormZC(x).Tag Then
FormZC(x).TBox1.NewLine tex
End If
p = True
End If
x = x + 1
Loop



End If
End Function
Public Function Adduser(ByVal tex As String, ByVal forma As String)

Do Until p
If formes(x) = forma Then
FormZC(x).nlist.AddItem tex
p = True
End If
x = x + 1
Loop
End Function
Public Function Remuser(ByVal tex As String, ByVal forma As String)

Do Until p
If formes(x) = forma Then
For g = 0 To FormZC(x).nlist.ListCount - 1
If FormZC(x).nlist.List(g) = tex Then
FormZC(x).nlist.RemoveItem g
g = FormZC(x).nlist.ListCount - 1
End If
Next g

p = True
End If
x = x + 1
Loop
End Function
