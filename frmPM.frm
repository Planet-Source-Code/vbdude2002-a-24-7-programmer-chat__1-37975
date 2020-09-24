VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPM 
   Caption         =   "Chat with "
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4935
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wsDirect 
      Left            =   600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsHost 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2400
      Top             =   1920
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   4935
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmPM.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu mnuDirect 
      Caption         =   "Direct"
      Begin VB.Menu mnuDirectCode 
         Caption         =   "Direct Code"
      End
      Begin VB.Menu mnuDirectSend 
         Caption         =   "Send File [Coming Soon]"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ClickRTB As Boolean
Private tmpCode As String
Public tmpCodefrm As New frmDirectCode
Private Sub Form_Unload(Cancel As Integer)
wsHost.Close
wsDirect.Close
End Sub


Private Sub mnuConnect_Click()
xsplit = Split(Caption, " ")
Randomize
rndNum = (Int(1000 * Rnd) + 1001) * 5
Dim xsp As String
xsp = xsplit(2)
Send MkMsg("DIRECTON", xsp, InputBox("Choose a port allowed through firewalls / routers", "Direct Connection", rndNum), " ")
End Sub


Private Sub mnuDirectCode_Click()
Dim D As New frmDirectCode
Load D
D.Tag = Me.Tag
D.Show , Me
End Sub

Private Sub mnuDisconnect_Click()
wsHost.Close
wsDirect.Close
End Sub

Private Sub rtbChat_GotFocus()
If ClickRTB = False Then txtSend.SetFocus
End Sub


Private Sub rtbChat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ClickRTB = True
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
txtSend.SetFocus
Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()
ClickRTB = False

If wsHost.State = sckConnected Or wsDirect.State = sckConnected Then
mnuDirect.Enabled = True
Else
mnuDirect.Enabled = False
End If

End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
Dim spupname As String
If KeyCode = 13 Then
    spup = Split(Me.Caption, " ")
    spupname = spup(2)
    Send MkMsg("PM", spupname, txtSend)
    rtbChat.SelLength = 0
    rtbChat.SelStart = Len(rtbChat.Text)
    rtbChat.SelColor = vbBlue
    rtbChat.SelText = "<" & Id.YourName & ">: "
    rtbChat.SelColor = vbBlack
    rtbChat.SelText = txtSend & vbCrLf
    txtSend = ""
End If

End Sub

Private Sub wsDirect_Close()
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Direct Connection Closed" & vbCrLf
rtbChat.SelBold = False
End Sub

Private Sub wsDirect_Connect()
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Directly Connected!" & vbCrLf
rtbChat.SelBold = False
End Sub

Private Sub wsDirect_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
wsDirect.GetData Data
datstr = Split(Data, "|%|")
i = 0
For i = LBound(datstr) To UBound(datstr)
Dat = Split(datstr(i), "|@|")


Select Case UCase(Dat(0))
Case "CODEDATA"
frmGetCode.Show , frmMain
frmGetCode.rtb.Text = frmGetCode.rtb.Text & Dat(1)
End Select
'wsDirect.SendData "CODEDATA|@|" & codesplit(i) & "|%|"
Next i
End Sub

Private Sub wsDirect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Direct Connection Error: " & Description & vbCrLf
rtbChat.SelBold = False
End Sub

Private Sub wsHost_Close()
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = Clr.clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Direct Connection Closed!" & vbCrLf
rtbChat.SelBold = False
End Sub

Private Sub wsHost_ConnectionRequest(ByVal requestID As Long)
wsHost.Close
wsHost.Accept requestID
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Directly Connected!" & vbCrLf
rtbChat.SelBold = False

End Sub

Private Sub wsHost_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
wsDirect.GetData Data
datstr = Split(Data, "|%|")
i = 0
For i = LBound(datstr) To UBound(datstr)
Dat = Split(datstr(i), "|@|")

Select Case UCase(Dat(0))
Case "CODEDATA"
frmGetCode.Show , frmMain
frmGetCode.rtb.Text = frmGetCode.rtb.Text & Dat(1)
End Select
'wsDirect.SendData "CODEDATA|@|" & codesplit(i) & "|%|"
Next i
End Sub

Private Sub wsHost_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
rtbChat.SelStart = Len(rtbChat.Text)
rtbChat.SelColor = clrGray
rtbChat.SelBold = True
rtbChat.SelText = vbCrLf & "Direct Connection Error: " & Description & vbCrLf
rtbChat.SelBold = False
End Sub


