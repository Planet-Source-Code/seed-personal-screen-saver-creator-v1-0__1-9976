VERSION 5.00
Begin VB.Form frmopt3final 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   5415
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmopt3final"
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
frmOpenPicForOpt3.Show
Me.Hide
End Sub

Private Sub Form_Click()
ShowCursor (bShow = False) 'show mouse
frmOpenPicForOpt3.Show
Me.Hide
End Sub

Private Sub Form_Load()
Me.Show
On Error GoTo erhr
'load picture that was selected
Image1.Picture = LoadPicture(frmOpenPicForOpt3.Dir1.Path + "\" + frmOpenPicForOpt3.File1.FileName)
If frmDir.opt1.Value = True Then GoTo rtl
If frmDir.opt2.Value = True Then GoTo ltr
If frmDir.opt3.Value = True Then GoTo down
If frmDir.opt4.Value = True Then GoTo up
If frmDir.opt5.Value = True Then GoTo diag1
If frmDir.opt6.Value = True Then GoTo diag2
If frmDir.opt7.Value = True Then GoTo diag3
If frmDir.opt8.Value = True Then GoTo diag4
'the following for statements are looped over and over
'and the if...then's are used to relocate the picture once it
'gets out of the screen!
rtl:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Left = Frame1.Left - 100
pause 0.01
If Frame1.Left < -6845 Then
Frame1.Left = Me.Width
End If
Next i
Exit Sub
ltr:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Left = Frame1.Left + 100
pause 0.01
If Frame1.Left > 11520 Then
Frame1.Left = -6845
End If
Next i
Exit Sub
down:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top + 100
pause 0.01
If Frame1.Top > 8880 Then
Frame1.Top = -6745
End If
Next i
Exit Sub
up:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top - 100
pause 0.01
If Frame1.Top < -6700 Then
Frame1.Top = 8880
End If
Next i
Exit Sub
diag1:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top + 100
Frame1.Left = Frame1.Left - 100
pause 0.01
If Frame1.Top > 8880 Then
Frame1.Top = -6700: Frame1.Left = Me.Width
End If
Next i
Exit Sub
diag2:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top + 100
Frame1.Left = Frame1.Left + 100
pause 0.01
If Frame1.Top > 8880 Then
Frame1.Top = -6700: Frame1.Left = -6845
End If
Next i
Exit Sub
diag3:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top - 100
Frame1.Left = Frame1.Left + 100
pause 0.01
If Frame1.Top < -6700 Then
Frame1.Top = 8880: Frame1.Left = -6845
End If
Next i
Exit Sub
diag4:
For i = 1 To 10000000
On Command1.Value GoTo exlad
Frame1.Top = Frame1.Top - 100
Frame1.Left = Frame1.Left - 100
pause 0.01
If Frame1.Top < -6700 Then
Frame1.Top = 8880: Frame1.Left = Me.Width
End If
Next i
Exit Sub
erhr:
MsgBox "An Error Occured.  Please check to see that the picture you chose was NOT directly in a drive, but rather in a folder.  This may have caused the error, and if not, then I don't know what to tell you!", vbCritical + vbOKOnly, "Error:"
Me.Hide
frmOpenPicForOpt3.Show
Exit Sub
exlad:
frmOpenPicForOpt3.Show
Me.Hide
End Sub

Private Sub Image1_Click()
ShowCursor (bShow = False) 'show mouse!
frmOpenPicForOpt3.Show
Me.Hide
End Sub
