VERSION 5.00
Begin VB.Form frmHMHL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose:"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmHMHL.frx":0000
      Left            =   3480
      List            =   "frmHMHL.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmHMHL.frx":002C
      Left            =   3480
      List            =   "frmHMHL.frx":004B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Choose how many pictures you want to add and also choose how long (in seconds) the delay time will be:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Number Of Pictures:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Delay Interval (in seconds):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "frmHMHL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1 = "" Or Combo2 = "" Then
MsgBox "Please Select a number from each box!", vbExclamation + vbOKOnly, "Error:"
Exit Sub
End If

Unload frmOpenPicsForOpt2
Load frmOpenPicsForOpt2
frmOpenPicsForOpt2.Show
Me.Hide

End Sub

Private Sub Command2_Click()
Unload Me
frmSelectPicBasedScrSaver.Show
Load Me
End Sub
