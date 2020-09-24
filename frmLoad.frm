VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frmLoad.frx":0000
      Top             =   0
      Width           =   1965
   End
   Begin VB.Image Image3 
      Height          =   2385
      Left            =   2040
      Picture         =   "frmLoad.frx":964E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Version 1.0b1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "PSCode VB Chat By:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   4095
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub loaddat()
frmMain.cbxServ.Clear
INI.INISetup App.Path & "/" & "Config.ini", 500
frmMain.cbxServ.Text = "Yazdi.no-ip.com"
frmMain.cbxServ.AddItem "Yazdi.no-ip.com"
frmMain.cbxServ.AddItem "Pscode.no-ip.org"
frmMain.cbxServ.AddItem "Vbchat.no-ip.org"
'cbxServ.AddItem "Vbchat.Zapto.org"
'cbxServ.AddItem "Vbchat.hopto.org"
'cbxServ.AddItem "68.4.160.70" '169.254.36.217
frmMain.cbxServ = "Pscode.no-ip.org"
frmEditProfile.txtName = INI.Read_Ini("Profile", "Name")
frmEditProfile.txtAge = INI.Read_Ini("Profile", "Age")
frmEditProfile.txtAbility = INI.Read_Ini("Profile", "Ability")
frmEditProfile.txtSite = INI.Read_Ini("Profile", "Website")
For i = 1 To 100
Load PM(i)
PM(i).Tag = i
Next i
Unload Me
frmMain.Show
End Sub
Private Sub Form_Load()
Label3 = Label3 & vbCrLf & "vbDude2002"
End Sub

Private Sub Timer1_Timer()
Call loaddat
Timer1.Enabled = False
End Sub
