VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1935
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1143.262
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbxClient 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Program type:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    On Error Resume Next
    Id.YourClient = "XXXX"
    Me.Hide
End Sub

Private Sub cmdOK_Click()

Id.YourName = txtUserName
Id.YourPass = txtPassword
Id.YourClient = Me.cbxClient
Select Case Id.YourClient
Case "Visual Basic"
    Id.YourClient = "MSVB"
Case "C++"
    Id.YourClient = "MSC+"
Case "Java"
    Id.YourClient = "JAVA"
Case "HTML"
    Id.YourClient = "HTML"
Case ".Net"
    Id.YourClient = "MS.N"
Case "Chat Client"
    Id.YourClient = "CHAT"
Case Else
    Id.YourClient = "CHAT"
    Me.Hide
    Exit Sub
End Select
Me.Hide
End Sub

Private Sub Form_Load()
cbxClient.AddItem "Visual Basic"
cbxClient.AddItem "C++"
cbxClient.AddItem "Java"
cbxClient.AddItem ".Net"
cbxClient.AddItem "HTML"
cbxClient.AddItem "Chat Client"
End Sub


