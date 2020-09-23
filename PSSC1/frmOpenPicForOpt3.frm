VERSION 5.00
Begin VB.Form frmOpenPicForOpt3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Picture:"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   1920
      Pattern         =   "*.bmp*;*.gif*;*.jpg*;*.wmf*"
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOpenPicForOpt3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'
ShowCursor (bShow = True) 'hide mouse
'
Unload frmopt3final
frmopt3final.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
frmDir.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo erhandler1
Dir1.Path = Drive1.Drive
Exit Sub
erhandler1:
MsgBox Err.Description, vbExclamation + vbOKOnly, "Error:"
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

