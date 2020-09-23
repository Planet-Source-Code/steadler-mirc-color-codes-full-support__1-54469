Attribute VB_Name = "Hooker"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''  S.O.S, The slightest error may cause VB to crash ''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim HookFRM As Form
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const PTR_WNDPROC       As Long = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Const PropName          As String = "Hooked"

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_ACTIVATE = &H6
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_ASKCBFORMATNAME = &H30C
Private Const WM_CANCELJOURNAL = &H4B
Private Const WM_CANCELMODE = &H1F
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_CHAR = &H102
Private Const WM_CHARTOITEM = &H2F
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_USER = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Private Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Private Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Private Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Private Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Private Const WM_PSD_MARGINRECT = (WM_USER + 3)
Private Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Private Const WM_PSD_PAGESETUPDLG = (WM_USER)
Private Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Private Const WM_CLEAR = &H303
Private Const WM_CLOSE = &H10
Private Const WM_COMMAND = &H111
Private Const WM_COMMNOTIFY = &H44
Private Const WM_COMPACTING = &H41
Private Const WM_COMPAREITEM = &H39
Private Const WM_CONVERTREQUESTEX = &H108
Private Const WM_COPY = &H301
Private Const WM_COPYDATA = &H4A
Private Const WM_CREATE = &H1
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_CTLCOLORMSGBOX = &H132
Private Const WM_CTLCOLORSCROLLBAR = &H137
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_CUT = &H300
Private Const WM_DDE_FIRST = &H3E0
Private Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Private Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Private Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Private Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Private Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Private Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Private Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Private Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Private Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Private Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Private Const WM_DEADCHAR = &H103
Private Const WM_DELETEITEM = &H2D
Private Const WM_DESTROY = &H2
Private Const WM_DESTROYCLIPBOARD = &H307
Private Const WM_DEVMODECHANGE = &H1B
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_DRAWITEM = &H2B
Private Const WM_DROPFILES = &H233
Private Const WM_ENABLE = &HA
Private Const WM_ENDSESSION = &H16
Private Const WM_ENTERIDLE = &H121
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_ERASEBKGND = &H14
Private Const WM_EXITMENULOOP = &H212
Private Const WM_FONTCHANGE = &H1D
Private Const WM_GETDLGCODE = &H87
Private Const WM_GETFONT = &H31
Private Const WM_GETHOTKEY = &H33
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_HOTKEY = &H312
Private Const WM_HSCROLL = &H114
Private Const WM_HSCROLLCLIPBOARD = &H30E
Private Const WM_ICONERASEBKGND = &H27
Private Const WM_IME_CHAR = &H286
Private Const WM_IME_COMPOSITION = &H10F
Private Const WM_IME_COMPOSITIONFULL = &H284
Private Const WM_IME_CONTROL = &H283
Private Const WM_IME_ENDCOMPOSITION = &H10E
Private Const WM_IME_KEYDOWN = &H290
Private Const WM_IME_KEYLAST = &H10F
Private Const WM_IME_KEYUP = &H291
Private Const WM_IME_NOTIFY = &H282
Private Const WM_IME_SELECT = &H285
Private Const WM_IME_SETCONTEXT = &H281
Private Const WM_IME_STARTCOMPOSITION = &H10D
Private Const WM_INITDIALOG = &H110
Private Const WM_INITMENU = &H116
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYLAST = &H108
Private Const WM_KEYUP = &H101
Private Const WM_KILLFOCUS = &H8
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDICASCADE = &H227
Private Const WM_MDICREATE = &H220
Private Const WM_MDIDESTROY = &H221
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_MDINEXT = &H224
Private Const WM_MDIREFRESHMENU = &H234
Private Const WM_MDIRESTORE = &H223
Private Const WM_MDISETMENU = &H230
Private Const WM_MDITILE = &H226
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUCHAR = &H120
Private Const WM_MENUSELECT = &H11F
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSELAST = &H209
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOVE = &H3
Private Const WM_NCACTIVATE = &H86
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCCREATE = &H81
Private Const WM_NCDESTROY = &H82
Private Const WM_NCHITTEST = &H84
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCMBUTTONDBLCLK = &HA9
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_NCPAINT = &H85
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5
Private Const WM_NEXTDLGCTL = &H28
Private Const WM_NULL = &H0
Private Const WM_OTHERWINDOWCREATED = &H42
Private Const WM_OTHERWINDOWDESTROYED = &H43
Private Const WM_PAINT = &HF
Private Const WM_PAINTCLIPBOARD = &H309
Private Const WM_PAINTICON = &H26
Private Const WM_PALETTECHANGED = &H311
Private Const WM_PALETTEISCHANGING = &H310
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_PASTE = &H302
Private Const WM_PENWINFIRST = &H380
Private Const WM_PENWINLAST = &H38F
Private Const WM_POWER = &H48
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_QUERYOPEN = &H13
Private Const WM_QUEUESYNC = &H23
Private Const WM_QUIT = &H12
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RENDERALLFORMATS = &H306
Private Const WM_RENDERFORMAT = &H305
Private Const WM_SETCURSOR = &H20
Private Const WM_SETFOCUS = &H7
Private Const WM_SETFONT = &H30
Private Const WM_SETHOTKEY = &H32
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_SHOWWINDOW = &H18
Private Const WM_SIZE = &H5
Private Const WM_SIZECLIPBOARD = &H30B
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_SYSCHAR = &H106
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_SYSCOMMAND = &H112
Private Const WM_SYSDEADCHAR = &H107
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_TIMECHANGE = &H1E
Private Const WM_TIMER = &H113
Private Const WM_UNDO = &H304
Private Const WM_VKEYTOITEM = &H2E
Private Const WM_VSCROLL = &H115
Private Const WM_VSCROLLCLIPBOARD = &H30A
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WININICHANGE = &H1A

Private HookhWnd      As Long
Private OriginalProcPointer As Long
Public No As Long
Public PNo As Long
Private Function Event_Proc(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  'Window procedure for all IDE MDI Child Windows
  'Intercepts all messages to code pane windows
    OriginalProcPointer = GetProp(hwnd, PropName) 'get original winproc pointer from this window's property
    If OriginalProcPointer Then 'if we got one then this window is subclassed:
        Event_Proc = CallWindowProc(OriginalProcPointer, hwnd, nMsg, wParam, lParam) 'call the original winproc to do what has to be done
        Select Case nMsg 'and now split on message type


           Case WM_MBUTTONDOWN
              MsgBox "Mouse Down"
           Case WM_MOUSEWHEEL
            
             If Form1.TBox1.sValue - (wParam / &H780000) < 1 Then
                 Form1.TBox1.sValue = 1
                 Form1.TBox1.UpdateBar
              ElseIf Form1.TBox1.sValue - (wParam / &H780000) > Form1.TBox1.sMax Then
                 Form1.TBox1.sValue = Form1.TBox1.sMax
                 Form1.TBox1.UpdateBar
             Else
                 Form1.TBox1.sValue = Form1.TBox1.sValue - (wParam / &H780000)
                 Form1.TBox1.UpdateBar
              End If
        End Select
            
     End If

End Function

Public Sub HookWindowEvents(hwnd&)
    HookhWnd = hwnd&
    If HookhWnd Then
       SetProp HookhWnd, PropName, GetWindowLong(HookhWnd, PTR_WNDPROC) 'store the winproc pointer of this window in a property
       SetWindowLong HookhWnd, PTR_WNDPROC, AddressOf Event_Proc 'and now point to our Proc so that we can see the messages arriving at this code pane
    End If
End Sub

Public Sub UnhookWindowEvents()
    OriginalProcPointer = GetProp(HookhWnd, PropName) 'get the original code pane winproc pointer from property
    If OriginalProcPointer Then 'if there is one then we unhook this window:
        RemoveProp HookhWnd, PropName 'remove the property
        SetWindowLong HookhWnd, PTR_WNDPROC, OriginalProcPointer 'and restore the original winproc pointer
    End If
    
    
End Sub

Function ClassName$(hwnd&)
Dim R&, sClassName$
  sClassName = String(100, 0)
  R = GetClassName(hwnd, sClassName, 100)
  ClassName = Left(sClassName, R)
End Function
':) Ulli's VB Code Formatter V2.15.4 (05.09.2002 17:52:16) 31 + 69 = 100 Lines

