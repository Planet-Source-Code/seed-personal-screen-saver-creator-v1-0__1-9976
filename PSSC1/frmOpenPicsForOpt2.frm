VERSION 5.00
Begin VB.Form frmOpenPicsForOpt2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Pictures:"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Pics"
      Height          =   375
      Left            =   2400
      TabIndex        =   27
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview:"
      Height          =   2175
      Left            =   4920
      TabIndex        =   25
      Top             =   360
      Width           =   2295
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6480
      Width           =   4335
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6120
      Width           =   4335
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5400
      Width           =   4335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5040
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Picture"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2400
      Pattern         =   "*.bmp*;*.gif*;*.jpg*;*.wmf*"
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
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
   Begin VB.Label Label12 
      Caption         =   "Pic 10:"
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Pic 9:"
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Pic 8:"
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Pic 7:"
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Pic 6:"
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Pic 5:"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Pic 4:"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Pic 3:"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Pic 2:"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Pic 1:"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "frmOpenPicsForOpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Private Sub Command1_Click()
On Error GoTo erhan

Image1.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)

If Me.Height = 4320 And Text1.Text <> "" Then Command3.Enabled = True
If Me.Height = 4680 And Text2.Text <> "" Then Command3.Enabled = True
If Me.Height = 5055 And Text3.Text <> "" Then Command3.Enabled = True
If Me.Height = 5415 And Text4.Text <> "" Then Command3.Enabled = True
If Me.Height = 5745 And Text5.Text <> "" Then Command3.Enabled = True
If Me.Height = 6135 And Text6.Text <> "" Then Command3.Enabled = True
If Me.Height = 6480 And Text7.Text <> "" Then Command3.Enabled = True
If Me.Height = 6825 And Text8.Text <> "" Then Command3.Enabled = True
If Me.Height = 7200 And Text9.Text <> "" Then Command3.Enabled = True

If File1.FileName = "" Then MsgBox "Select a file, not just a directory or a drive!", vbExclamation + vbOKOnly, "Error:": Exit Sub

If Me.Height = 4320 Then GoTo jtwo
If Me.Height = 4680 Then GoTo jthree
If Me.Height = 5055 Then GoTo jfour
If Me.Height = 5415 Then GoTo jfive
If Me.Height = 5745 Then GoTo jsix
If Me.Height = 6135 Then GoTo jseven
If Me.Height = 6480 Then GoTo jeight
If Me.Height = 6825 Then GoTo jnine
If Me.Height = 7200 Then GoTo jten

jtwo:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jthree:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jfour:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jfive:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jsix:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text6.Text = "" Then Text6.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jseven:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text6.Text = "" Then Text6.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text7.Text = "" Then Text7.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jeight:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text6.Text = "" Then Text6.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text7.Text = "" Then Text7.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text8.Text = "" Then Text8.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jnine:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text6.Text = "" Then Text6.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text7.Text = "" Then Text7.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text8.Text = "" Then Text8.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text9.Text = "" Then Text9.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub
jten:
If Text1.Text = "" Then Text1.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text2.Text = "" Then Text2.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text3.Text = "" Then Text3.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text4.Text = "" Then Text4.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text5.Text = "" Then Text5.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text6.Text = "" Then Text6.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text7.Text = "" Then Text7.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text8.Text = "" Then Text8.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text9.Text = "" Then Text9.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
If Text10.Text = "" Then Text10.Text = Dir1.Path + "\" + File1.FileName: Exit Sub
Exit Sub

erhan:
MsgBox Err.Description, vbCritical + vbOKOnly, "Error:"
MsgBox "If the picture you tried to load was directly in one of your drives, then please move that file to a FOLDER.  I'm sorry about this small problem!", vbExclamation + vbOKOnly, "Note:"
Exit Sub
End Sub

Private Sub Command2_Click()
frmHMHL.Show
Unload Me
Load Me
End Sub

Private Sub Command3_Click()
'
ShowCursor (bShow = True)
'
Me.Hide
Load frmOpt2Final
frmOpt2Final.Show
'
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo erh
Dir1.Path = Drive1.Drive
GoTo done
erh:
MsgBox Err.Description, vbCritical + vbOKOnly, "Error:"
Exit Sub
done:
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

Private Sub Form_Load()

hm = frmHMHL.Combo1

If hm = 2 Then Me.Height = 4320: Exit Sub
If hm = 3 Then Me.Height = 4680: Exit Sub
If hm = 4 Then Me.Height = 5055: Exit Sub
If hm = 5 Then Me.Height = 5415: Exit Sub
If hm = 6 Then Me.Height = 5745: Exit Sub
If hm = 7 Then Me.Height = 6135: Exit Sub
If hm = 8 Then Me.Height = 6480: Exit Sub
If hm = 9 Then Me.Height = 6825: Exit Sub
If hm = 10 Then Me.Height = 7200: Exit Sub

End Sub
