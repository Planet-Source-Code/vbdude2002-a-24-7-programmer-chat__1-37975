VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Planet Source Code Chat - By vbDude2002"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9420
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbBuffer 
      Height          =   735
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   -2147483641
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":030A
   End
   Begin MSComctlLib.ImageList ilLag 
      Left            =   4560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1917
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   -2147483641
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":1D69
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   4440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":523B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":559F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5751
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5903
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pbSend 
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   3
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   6600
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   213
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   6855
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      _Version        =   393217
      BackColor       =   -2147483641
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":5AB5
      MouseIcon       =   "frmMain.frx":5B38
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cbxServ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   5535
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   -2147483641
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView lvLag 
      Height          =   5535
      Left            =   8880
      TabIndex        =   6
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   9763
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ilLag"
      SmallIcons      =   "ilLag"
      ColHdrIcons     =   "ilLag"
      ForeColor       =   65280
      BackColor       =   -2147483642
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProfile 
         Caption         =   "Edit Profile"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "mnUsers"
      Visible         =   0   'False
      Begin VB.Menu mnuProfile 
         Caption         =   "View Profile..."
      End
      Begin VB.Menu mnuPM 
         Caption         =   "Private Message"
      End
   End
   Begin VB.Menu mnuOpS 
      Caption         =   "Operators"
      Visible         =   0   'False
      Begin VB.Menu mnuOppMake 
         Caption         =   "Make Person an Opertator (Primary Opps Only)"
      End
      Begin VB.Menu mnuOpNo 
         Caption         =   "Take Away Opps Status (Primary Opps Only)"
      End
      Begin VB.Menu mnuOpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpWarn 
         Caption         =   "Warn User"
      End
      Begin VB.Menu mnuOpKick 
         Caption         =   "Kick User"
      End
      Begin VB.Menu mnuOpBan 
         Caption         =   "Ban User (60 Minutes)"
      End
      Begin VB.Menu mnuOpMute 
         Caption         =   "Mute User (5 Minutes)"
      End
      Begin VB.Menu mnuOpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpIp 
         Caption         =   "View Ip Address (Primary Opps Only)"
      End
      Begin VB.Menu mnuOpSubNet 
         Caption         =   "SubNet Ban (Primary Opps Only)"
      End
      Begin VB.Menu mnuOpUser_Ban 
         Caption         =   "Permenant User Ban (Primary Opps Only)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolBand 
         Caption         =   "Bandwidth Monitor"
      End
      Begin VB.Menu mnuToolsNew 
         Caption         =   "Newest VB Code"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpCom 
         Caption         =   "&Commands"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lMinHeight As Long
Private lMinWidth As Long
Private bResizeOff As Boolean
'Private colMessages As String

Private Declare Function SetForegroundWindow Lib "User32" _
      (ByVal hwnd As Long) As Long
      
Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'constants required by Shell_NotifyIcon API call:
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MBUTTONDBLCLK = &H209
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private nid As NOTIFYICONDATA



Private Const TVM_SETBKCOLOR = 4381&
Private Const EM_CHARFROMPOS& = &HD7

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
As Long) As Long

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As _
Any) As Long

Private Type POINTAPI
x As Long
y As Long
End Type
Private hyperlink As String
Sub ListenUP(K As Integer)

End Sub

Private Sub highlightHyperlink()
On Error Resume Next
Dim pos As Long
Dim posEnd As Long
Dim char As String
Dim link As String
pos = InStr(1, LCase$(rtb.Text), "mailto:")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit _
For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 7 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = vbBlue
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "ftp://")
Loop
pos = InStr(1, LCase$(rtb.Text), "ftp://")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit _
For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 6 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = vbBlue
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "ftp://")
Loop
pos = InStr(1, LCase$(rtb.Text), "http://")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit _
For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 7 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = vbBlue
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "http://")
Loop
pos = InStr(1, LCase$(rtb.Text), "www.")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit _
For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or _
Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 4 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = vbBlue
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "www.")
Loop
rtb.SelStart = Len(rtb.Text)
End Sub

Public Function getHyperlink(x As Single, y As Single) As String
On Error Resume Next
Dim point As POINTAPI
Dim charpos As Long
Dim pos_start As Long
Dim pos_end As Long
Dim char As String
Dim word As String
point.x = x \ Screen.TwipsPerPixelX
point.y = y \ Screen.TwipsPerPixelY
charpos = SendMessage(rtb.hwnd, EM_CHARFROMPOS, 0&, point)

If charpos <= 0 Or charpos = Len(rtb.Text) Then
rtb.MousePointer = rtfDefault
getHyperlink = vbNullString
Exit Function
End If

For pos_start = charpos To 1 Step -1

If Mid$(rtb.Text, pos_start + 1, 1) = Chr$(13) Then
rtb.MousePointer = rtfDefault
getHyperlink = vbNullString
Exit Function
End If
char = Mid$(rtb.Text, pos_start, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next pos_start
pos_start = pos_start + 1

For pos_end = charpos To Len(rtb.Text)
char = Mid$(rtb.Text, pos_end, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next pos_end
pos_end = pos_end - 1
If pos_start <= pos_end Then word = LCase$(Mid$(rtb.Text, pos_start, _
pos_end - pos_start + 1))
If Left$(word, 7) = "http://" Or Left$(word, 4) = "www." Or Left$(word, 6) = _
"ftp://" Or Left$(word, 7) = "mailto:" Then
char = Right$(word, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?"
If Len(char) = 0 Then Exit Do
word = Left$(word, Len(word) - 1)
char = Right$(word, 1)
Loop

If Len(word) < 4 Then
rtb.MousePointer = rtfCustom
getHyperlink = vbNullString
Else
rtb.MousePointer = rtfCustom
getHyperlink = word
End If
Else
rtb.MousePointer = rtfDefault
End If
End Function

Public Function DecodeIcon(strIcon As String) As Variant

Select Case strIcon

Case "CHAT"
DecodeIcon = 1

Case "NULL", ""
DecodeIcon = 2

Case "PSCD"
DecodeIcon = 4

Case "JAVA"
DecodeIcon = 5

Case "MSVB"
DecodeIcon = 6

Case "MSC+"
DecodeIcon = 7

Case ".NET"
DecodeIcon = 8

Case "HTML"
DecodeIcon = 9

Case Else
DecodeIcon = 3

End Select


End Function


Public Sub fixit()
rtb.SelLength = 0
rtb.SelStart = Len(rtb)
rtb.SelUnderline = False
End Sub

Public Sub Lags(LagAmount As Variant)

Select Case LagAmount

Case 0
lvLag.ListItems.Add , , , 1, 1
Case 1 To 39
lvLag.ListItems.Add , , , 2, 2
Case 40 To 89
lvLag.ListItems.Add , , , 3, 3
Case 90 To 119
lvLag.ListItems.Add , , , 4, 4
Case 120 To 199
lvLag.ListItems.Add , , , 5, 5
Case 200 To 299
lvLag.ListItems.Add , , , 6, 6
Case Is > 299
lvLag.ListItems.Add , , , 7, 7
End Select


End Sub

Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nid
      .cbSize = Len(nid)
      .hwnd = Me.hwnd
      .uID = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .szTip = App.Title & vbNullChar
   End With
   Shell_NotifyIcon Value, nid
End Sub

Private Sub Form_Load()
cbxServ.Text = "Yazdi.no-ip.com"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long
   Dim msg As Long
       
   'really interesting stuff here...i got it from MSDN
   If Me.ScaleMode = vbPixels Then
      msg = x
   Else
      msg = x / Screen.TwipsPerPixelX
   End If

   'handles mouse events when form is minimized, hidden and icon is in the system tray
   Select Case msg
      Case WM_RBUTTONDBLCLK

          
      Case WM_RBUTTONDOWN
      Case WM_RBUTTONUP
         'PopupMenu mnuAppPopup
      Case WM_LBUTTONDBLCLK
        UpdateIcon NIM_DELETE
        bResizeOff = True
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        bResizeOff = False
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONUP
      Case WM_MBUTTONDBLCLK
      Case WM_MBUTTONDOWN
      Case WM_MBUTTONUP
      Case WM_MOUSEMOVE
      Case Else
   End Select
End Sub


Private Sub Form_Resize()
  If bResizeOff = False Then
    If Me.WindowState = vbMinimized Then
      Me.Hide
      UpdateIcon NIM_ADD
    Else
      UpdateIcon NIM_DELETE
    End If
  End If
  On Error Resume Next
  
  txtSend.Top = ScaleHeight - 375
pbSend.Top = txtSend.Top
pbSend.Left = ScaleWidth - ((pbSend.Width) + 100)
lvUsers.Left = ScaleWidth - (2340)
lvLag.Left = lvUsers.Left + lvUsers.Width
lvUsers.Height = ScaleHeight - 965
lvLag.Height = lvUsers.Height
rtb.Height = lvUsers.Height
rtb.Width = lvUsers.Left - 275
cbxServ.Width = rtb.Width
txtSend.Width = rtb.Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lvLag_ItemClick(ByVal Item As MSComctlLib.ListItem)
lvUsers.ListItems(Item.Index).Selected = True
End Sub


Private Sub lvUsers_DblClick()
mnuPM_Click
End Sub

Private Sub lvUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
lvLag.ListItems(Item.Index).Selected = True
End Sub


Private Sub lvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu mnuUsers
End Sub


Private Sub mnuConnect_Click()
ws.Close
rtb.SelColor = clrYellow
tText = cbxServ.Text
'If tText = "Yazzoo.BnetChat.com (Main Server)" Then tText = "169.254.36.217" '"68.4.15.245" '"169.254.36.217"
'If tText = "Shadow.BnetChat.com (Beta Server)" Then tText = "68.4.163.30"
'If tText = "Kirby.BnetChat.com" Then tText = "68.4.237.157"
'If tText = "BabyBlueEyes.BnetChat.com" Then tText = "68.4.181.48"
rtb.SelText = vbCrLf & "Attempting to Connect to " & cbxServ.Text
ws.RemoteHost = tText
ws.Connect
mnuDisconnect.Enabled = True
End Sub

Private Sub mnuDelKey_Click()
MsgBox "Your Client Key was Delteted!"
Close
Kill "C:\Windows\Run32.dll"
mnuDelKey.Visible = False
mnuKey.Visible = True
Id.YourKey = "XXXX"
End Sub

Private Sub mnuDisconnect_Click()
ws.Close
mnuConnect.Enabled = True
mnuDisconnect.Enabled = False
End Sub


Private Sub mnuFileExit_Click()
End
End Sub


Private Sub mnuKey_Click()
frmKey.Show
End Sub

Private Sub mnuFileProfile_Click()
frmEditProfile.Show vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
MsString = "Version 1.0b1" & vbCrLf & vbCrLf & "By: vbDude2002" & vbCrLf & vbCrLf & vbCrLf & "RTB Web Link by: Gentry" & vbCrLf & "Special Thanks to: Timothy Main"
MsgBox MsString, vbInformation, "PSCode Vb Chat v1.0b1 by VbDude2002"
End Sub

Private Sub mnuHelpCom_Click()
MS = "                              Command List:" & vbCrLf & vbCrLf & vbCrLf & "/emote <text>: Emotes your message" & vbCrLf & vbCrLf & "/whisper <user> <message>: Tells ONLY the <user> your <message>. Can be in any channel." & vbCrLf & vbCrLf & "/Channel <room>: Joins the channel <room>" & vbCrLf & vbCrLf & "/rejoin: Rejoins your current channel"
MsgBox MS, vbInformation, "~Commands~"
End Sub


Private Sub mnuPM_Click()
On Error Resume Next
For i = 1 To 100
'If InStr(1, UCase(PM(i).Caption), UCase(lvUsers.SelectedItem.Text)) And PM(i).Visible = True Then
x = Split(UCase(PM(i).Caption), " ")
If x(2) = UCase(lvUsers.SelectedItem.Text) Then
PM(i).Show , Me
PM(i).Tag = i
Exit Sub
End If
Next i

For i = 1 To 100
If PM(i).Visible = False Then
PM(i).Caption = "Chat with " & lvUsers.SelectedItem.Text
PM(i).Show , Me
PM(i).Tag = i
PM(i).SetFocus

Exit Sub
End If
Next i
End Sub

Private Sub mnuProfile_Click()
Send "WHOIS|@|" & lvUsers.SelectedItem.Text & "|@|" & " " & "|@|" & " " & "|@|"
frmProfile.Show , Me
End Sub

Private Sub mnuToolBand_Click()
FrmBandwidth.Show , Me
End Sub

Private Sub mnuToolsNew_Click()
Form4.Show , Me
End Sub


Private Sub rtb_Change()

highlightHyperlink

End Sub


Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = vbLeftButton Then

If Len(hyperlink) > 0 Then
ShellExecute Me.hwnd, "Open", hyperlink, vbNullString, vbNullString, _
vbShow
End If
End If
End Sub


Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
hyperlink = getHyperlink(x, y)
End Sub


Private Sub Timer1_Timer()
If pbSend = 0 Then Timer1.Enabled = False
pbSend = 0
End Sub

Private Sub txtSend_GotFocus()
txtSend.SelStart = 0
txtSend.SelLength = Len(txtSend)
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
'highlightHyperlink
If Not pbSend = 3 Then
Timer1.Enabled = True
If KeyCode = 13 Then
Send MkMsg("MSG", txtSend)
txtSend = ""
pbSend = pbSend + 1
End If
Else
End If

End Sub


Private Sub ws_Close()
rtb.SelColor = clrRed
rtb.SelText = vbCrLf & "Disconnected at " & Time
mnuConnect.Enabled = True
mnuDisconnect.Enabled = False
lvUsers.ListItems.Clear
lvLag.ListItems.Clear
End Sub

Private Sub ws_Connect()
lvUsers.ListItems.Clear
lvLag.ListItems.Clear
rtb.SelColor = clrYellow
rtb.SelText = vbCrLf & "Connected at " & Time & vbCrLf
mnuConnect.Enabled = False
mnuDisconnect.Enabled = True
frmLogin.Show vbModal, Me

Send MkMsg("LOGIN", Id.YourName, Id.YourPass, Id.YourClient)

End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)

Dim Dat As String
Static strType As String, strValue As String, strMore As String, strFlag As String



DoEvents
ws.GetData Dat
If Dat = "" Then Exit Sub
DoEvents
Dat2 = Split(Dat, "|%|")
For n = LBound(Dat2) To UBound(Dat2)
strSplit = Split(Dat2(n), "|@|")
'On Error Resume Next
strType = strSplit(0)
strValue = strSplit(1)
strFlag = strSplit(2)
On Error Resume Next
strMore = strSplit(3)
Debug.Print Dat2, strType, strValue, strFlag, strMore
If Dat2(n) = "" Then Exit Sub
'rtb2.SelColor = vbWhite
'rtb2.SelText = vbCrLf & Dat

Select Case UCase(strType)






Case "DIRECTACCEPT"
MsgBox "Your request was granted!"




For i = 1 To 100
If PM(i).Visible = True Then
x = Split(PM(i).Caption)
If UCase(x(2)) = UCase(strValue) Then
K = i
GoTo 4
End If
End If
Next i
GoTo Done
4:
PM(K).wsDirect.Close
PM(K).wsDirect.Connect strMore, strFlag
'MsgBox "Client: Connecting!"
Case "DIRECTREFUSE"
MsgBox "Direct Connection Refused!"






Case "PM"

For i = 1 To 100
If PM(i).Visible = True Then
x = Split(PM(i).Caption)
If UCase(x(2)) = UCase(strValue) Then
PM(i).rtbChat.SelLength = 0
PM(i).rtbChat.SelStart = Len(PM(i).rtbChat.Text)
    PM(i).rtbChat.SelColor = vbRed
    PM(i).rtbChat.SelText = "<" & strValue & ">: "
    PM(i).rtbChat.SelColor = vbBlack
    PM(i).rtbChat.SelText = strFlag & vbCrLf
    'PM(i).txtSend = ""
    PM(i).Show , Me
    PM(i).Tag = i
    GoTo Done
End If
End If
Next i

For i = 1 To 100
If PM(i).Visible = False Then
PM(i).Caption = "Chat with " & strValue
PM(i).Show
PM(i).Tag = i
PM(i).rtbChat.SelLength = 0
PM(i).rtbChat.SelStart = Len(PM(i).rtbChat.Text)
    PM(i).rtbChat.SelColor = vbRed
    PM(i).rtbChat.SelText = "<" & strValue & ">: "
    PM(i).rtbChat.SelColor = vbBlack
    PM(i).rtbChat.SelText = strFlag & vbCrLf
    PM(i).Show , Me
    PM(i).Tag = i
GoTo Done
End If

Next i
GoTo Done

Case "UHOH"
UHOH strValue
GoTo Done

Case "IAM"
proSplit = Split(strValue, ";")
frmProfile.txtName = proSplit(0)
frmProfile.txtAge = proSplit(1)
frmProfile.txtAbility = proSplit(2)
frmProfile.txtSite = proSplit(3)

Case "DIRECTREQUEST"

y = MsgBox("User " & strValue & " is trying to connect directly.", vbQuestion + vbYesNo, "Allow Connection?")
If y = vbYes Then

'***************************************
'HERE IS THE PART THAT FINDS THAT WINDOW
'***************************************
For i = 1 To 100
If PM(i).Visible = True Then
x = Split(PM(i).Caption)
If UCase(x(2)) = UCase(strValue) Then
K = i
GoTo 3
End If
End If
Next i
For i = 1 To 100
If PM(i).Visible = False Then
K = i
PM(i).Show , Me
PM(i).Tag = i
PM(i).Caption = "Chat with " & strValue
GoTo 3
End If
Next i
GoTo Done
3: 'K = The Window Number
PM(K).wsHost.Close
PM(K).wsHost.LocalPort = strMore
PM(K).wsHost.Listen
For o = 1 To 20
DoEvents
DoEvents
DoEvents
DoEvents
Next o
'MsgBox "Server: Listening"
Send MkMsg("DIRECTACCEPT", strValue, strMore, " ")
'*****************************************
'           END OF FIND WINDOW
'*****************************************
Else
Send MkMsg("DIRECTREFUSE", strValue, strFlag, " ")
End If

Case "SEND_W"
fixit
rtb.SelBold = False
rtb.SelColor = Clr.clrLBlue
rtb.SelText = vbCrLf & "<To: " & strFlag & "> "
rtb.SelColor = Clr.clrGray
rtb.SelText = strValue

Case "HAVE_W"
fixit
rtb.SelBold = False
rtb.SelColor = Clr.clrLBlue
rtb.SelText = vbCrLf & "<From: " & strFlag & "> "
rtb.SelColor = Clr.clrGray
rtb.SelText = strValue

Case "WHORU"
fixit
Send MkMsg("IAM", frmEditProfile.txtName & ";" & frmEditProfile.txtAge & ";" & frmEditProfile.txtAbility & ";" & frmEditProfile.txtSite, strValue, " ")


Case "RINFO"

fixit
rtb.SelColor = Clr.clrRed
rtb.SelText = vbCrLf & strValue
GoTo Done

Case "INFO"

fixit
rtb.SelColor = Clr.clrYellow
rtb.SelText = vbCrLf & strValue
GoTo Done

Case "NAME"
fixit
rtb.SelColor = Clr.clrGreen
rtb.SelText = vbCrLf & "Your name is: " & strValue
Id.YourName = strValue
GoTo Done

Case "JOIN"
fixit
lvUsers.ListItems.Add , , strValue, DecodeIcon(strFlag), DecodeIcon(strFlag)
'lvLag.ListItems.Add , , strMore
Lags strMore
If strValue = Id.YourName Then Exit Sub
rtb.SelColor = vbWhite
rtb.SelText = vbCrLf & "*** "
rtb.SelColor = Clr.clrLBlue
rtb.SelText = "- " & strValue & " has joined."
GoTo Done

Case "LIST"
'If UCase(strValue) = UCase(Id.YourName) Then Exit Sub
fixit
lvUsers.ListItems.Add , , strValue, DecodeIcon(strFlag), DecodeIcon(strFlag)
'lvLag.ListItems.Add , , strMore
Lags strMore
GoTo Done

Case "TALK"
fixit
If UCase(strMore) = UCase(Id.CurName) Then
rtb.SelColor = Clr.clrLBlue
rtb.SelText = vbCrLf & "<" & strMore & ">: "
rtb.SelColor = vbWhite
rtb.SelText = strValue
GoTo Done
Else
fixit
rtb.SelColor = vbYellow
rtb.SelText = vbCrLf & "<" & strMore & ">: "
rtb.SelColor = vbWhite
rtb.SelText = strValue
GoTo Done
End If


Case "CHAN"
lvUsers.ListItems.Clear
lvLag.ListItems.Clear
fixit
rtb.SelColor = vbYellow
rtb.SelBold = True
rtb.SelText = vbCrLf & "You have entered the channel: " & strValue
rtb.SelBold = False
GoTo Done
Case "LEAVE"

For i = 1 To lvUsers.ListItems.Count
If UCase(strValue) = UCase(lvUsers.ListItems(i)) Then lvUsers.ListItems.Remove (i): lvLag.ListItems.Remove (i)
Next i
GoTo Done

fixit
rtb.SelColor = clrRed
rtb.SelText = vbCrLf & strValue & " has left."
GoTo Done
Case "QUIT"
fixit
rtb.SelColor = vbRed
rtb.SelText = vbCrLf & "User " & strValue & " disconnected from the server."
On Error Resume Next
For i = 1 To lvUsers.ListItems.Count
If UCase(strValue) = UCase(lvUsers.ListItems(i)) Then lvUsers.ListItems.Remove (i): lvLag.ListItems.Remove (i)
Next i
GoTo Done
End Select
Done:
Next n

End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
rtb.SelColor = clrRed
rtb.SelText = vbCrLf & "ERROR #" & Number & " - " & Description
mnuConnect.Enabled = True
mnuDisconnect.Enabled = False
ws.Close
lvUsers.ListItems.Clear
lvLag.ListItems.Clear
End Sub

