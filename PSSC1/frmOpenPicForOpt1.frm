VERSION 5.00
Begin VB.Form frmOpenPicForOpt1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Picture:"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Picture"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   2400
      Pattern         =   "*.bmp*;*.gif*;*.jpg*;*.wmf*"
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmOpenPicForOpt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'
ShowCursor (bShow = True) 'hide mouse pointer
'
On Error GoTo er67h 'error handler
'
frmOpt1Final.Image1.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName) 'loads selected pic in the next forms image box
frmOpt1Final.Show
Me.Hide
Exit Sub
er67h:
MsgBox Err.Description, vbOKOnly + vbCritical, "Error:"
MsgBox "If the file you tried to load was in a drive and not a folder, then you must move that file to a folder and try again.  I'm sorry about this small problem!", vbExclamation + vbOKOnly, "Note:"
Exit Sub

End Sub

Private Sub Command2_Click()
frmSelectPicBasedScrSaver.Show
Unload Me
Load Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo erh 'error handler
Dir1.Path = Drive1.Drive
Exit Sub
erh:
MsgBox Err.Description, vbCritical + vbOKOnly, "Error:"
Exit Sub
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub
