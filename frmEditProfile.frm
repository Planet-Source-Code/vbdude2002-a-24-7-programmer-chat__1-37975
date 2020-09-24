VERSION 5.00
Begin VB.Form frmEditProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Profile..."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmEditProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtSite 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtAbility 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Website:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Age:"
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Specialty:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmEditProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
INI.Write_Ini "Profile", "Name", txtName
INI.Write_Ini "Profile", "Age", txtAge
INI.Write_Ini "Profile", "Ability", txtAbility.Text
INI.Write_Ini "Profile", "Website", txtSite.Text
DoEvents
frmEditProfile.txtName = INI.Read_Ini("Profile", "Name")
frmEditProfile.txtAge = INI.Read_Ini("Profile", "Age")
frmEditProfile.txtAbility = INI.Read_Ini("Profile", "Ability")
frmEditProfile.txtSite = INI.Read_Ini("Profile", "Website")
Hide
End Sub

Private Sub Command2_Click()
Hide
frmEditProfile.txtName = INI.Read_Ini("Profile", "Name")
frmEditProfile.txtAge = INI.Read_Ini("Profile", "Age")
frmEditProfile.txtAbility = INI.Read_Ini("Profile", "Ability")
frmEditProfile.txtSite = INI.Read_Ini("Profile", "Website")
End Sub

Private Sub Form_Load()
frmEditProfile.txtName = INI.Read_Ini("Profile", "Name")
frmEditProfile.txtAge = INI.Read_Ini("Profile", "Age")
frmEditProfile.txtAbility = INI.Read_Ini("Profile", "Ability")
frmEditProfile.txtSite = INI.Read_Ini("Profile", "Website")
End Sub
