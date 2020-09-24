VERSION 5.00
Begin VB.Form FrmBandwidth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bandwidth Monitor v1.0b1"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBandwidth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   0
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   915
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   915
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   75
      X2              =   2115
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   75
      X2              =   2115
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   2115
      X2              =   2115
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   2115
      X2              =   3795
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      X1              =   3795
      X2              =   3795
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line Line7 
      X1              =   75
      X2              =   3795
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line8 
      X1              =   75
      X2              =   3795
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   915
      X2              =   915
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3315
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3315
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "FrmBandwidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single
Private Sub Form_Load()
' Global SentMax As Long, RecMax As Long

On Error Resume Next
Static a As String
Static b As String

Me.Top = 0
Me.Left = 0
Set m_objIpHelper = New CIpHelper
DoEvents
'frmMain.Timer1.Enabled = False
Call UpdateInterfaceInfo
DoEvents
'Me.Timer1.Enabled = True
'Me.Icon = frmMain.ImageList2.ListImages(1).Picture
'frmMain.StatusBar1.Panels(4).Picture = frmMain.ImageList2.ListImages(1).Picture
DoEvents
Label5.Caption = Me.lblRecv.Caption
Label6.Caption = Me.lblSent.Caption
DoEvents
Timer2.Enabled = True
On Error Resume Next
pbSent.Max = General.SentMax
pbRec.Max = General.RecMax
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.Timer1.Enabled = False
'frmMain.Timer1.Enabled = True
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub Label10_Click()
End Sub

Private Sub Timer1_Timer()
Call UpdateInterfaceInfo
End Sub
Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
End Sub

Private Sub Timer2_Timer()
'On Error Resume Next
DoEvents
Dim XX As Long
Dim YY As Long
Dim XXX As Long
Dim YYY As Long
DoEvents
XX = Me.lblRecv.Caption - YY




Static tR As Long, tS As Long

tR = Me.lblRecv.Caption
tS = Me.lblSent.Caption
'End With

DoEvents
TransferRate = Format(Int(XX) / 1024, "####.00")
DoEvents
TransferRate2 = Format(Int(XXX) / 1024, "####.00")
DoEvents
 
        


End Sub

