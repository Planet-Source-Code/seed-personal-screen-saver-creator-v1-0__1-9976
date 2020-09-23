VERSION 5.00
Begin VB.Form frmSelectPicBasedScrSaver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Type..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Several Alternating Pictures"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Single Moving Picture"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Single, Non-Moving Picture"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Please select what type of picture-based screen saver you would like:"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmSelectPicBasedScrSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
 If Option1.Value = True Then GoTo opt1
 If Option3.Value = True Then GoTo opt3
 If Option2.Value = True Then GoTo opt2

Exit Sub
opt1:
 frmOpenPicForOpt1.Show
 Unload Me
 Exit Sub
opt3:
 frmHMHL.Show
 Unload Me
 Exit Sub
opt2:
 frmDir.Show
 Unload Me
 Exit Sub
'
End Sub

Private Sub Command2_Click()
Unload Me
frmMain.Show
Load Me
End Sub
