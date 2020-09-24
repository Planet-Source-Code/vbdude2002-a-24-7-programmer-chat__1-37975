VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDirectCode 
   Caption         =   "Direct Code Transfer"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "frmDirectCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Progress 
      BackColor       =   &H80000009&
      Caption         =   "Progress..."
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   5175
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0 Bytes Recieved"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0 Bytes Sent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Code"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbCode 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10610
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDirectCode.frx":1042
   End
End
Attribute VB_Name = "frmDirectCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Progress.Visible = True
Label1 = "0/0 Bytes Sent"
Label2 = "0/0 Bytes Recieved"
pb.Value = 0
On Error Resume Next
codesplit = Split(rtbCode.Text, vbCrLf)
pb.Max = UBound(codesplit) + 1

If PM(Int(Tag)).wsDirect.State = sckConnected Then
        'PM(Int(Tag)).wsDirect.SendData "CODELEN|@|" & pb.Max & "|%|"
    For i = LBound(codesplit) To UBound(codesplit)
        PM(Int(Tag)).wsDirect.SendData "CODEDATA|@|" & codesplit(i) & vbCrLf & "|%|"
        DoEvents
        DoEvents
        DoEvents
        pb = pb + 1
        If Not i = UBound(codesplit) Then pb = pb + 1
    Next i
    'MsgBox "Code Sent..."
    Progress.Visible = False
End If

If PM(Int(Tag)).wsHost.State = sckConnected Then
        'PM(Int(Tag)).wsHost.SendData "CODELEN|@|" & pb.Max & "|%|"
    For i = LBound(codesplit) To UBound(codesplit)
        PM(Int(Tag)).wsHost.SendData "CODEDATA|@|" & codesplit(i) & vbCrLf & "|%|"
        DoEvents
        DoEvents
        DoEvents
        pb = pb + 1
        If Not i = UBound(codesplit) Then pb = pb + 1
    Next i
    'MsgBox "Code Sent..."
    Progress.Visible = False
End If

End Sub
