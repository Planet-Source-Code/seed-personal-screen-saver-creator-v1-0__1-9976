VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Direction:"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt8 
      Caption         =   "Diagonal (Bottom Right To Top Left)"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   3015
   End
   Begin VB.OptionButton opt7 
      Caption         =   "Diagonal (Bottom Left To Top Right)"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton opt6 
      Caption         =   "Diagonal (Top Left To Bottom Right)"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Left To Right"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton opt4 
      Caption         =   "Up"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton opt5 
      Caption         =   "Diagonal (Top Right To Bottom Left)"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton opt3 
      Caption         =   "Down"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Right To Left"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmOpenPicForOpt3.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
frmSelectPicBasedScrSaver.Show
End Sub
