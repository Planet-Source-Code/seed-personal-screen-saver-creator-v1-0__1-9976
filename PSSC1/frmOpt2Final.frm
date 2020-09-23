VERSION 5.00
Begin VB.Form frmOpt2Final 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmOpt2Final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub pause(interval) 'pause function
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub


Private Sub Form_Click()
'
ShowCursor (bShow = False) 'show mouse
'
Me.Hide
frmHMHL.Show
Unload frmOpenPicsForOpt2
'
frmOpenPicsForOpt2.Text1.Text = "" 'reset all boxes on last form!
frmOpenPicsForOpt2.Text2.Text = ""
frmOpenPicsForOpt2.Text3.Text = ""
frmOpenPicsForOpt2.Text4.Text = ""
frmOpenPicsForOpt2.Text5.Text = ""
frmOpenPicsForOpt2.Text6.Text = ""
frmOpenPicsForOpt2.Text7.Text = ""
frmOpenPicsForOpt2.Text8.Text = ""
frmOpenPicsForOpt2.Text9.Text = ""
frmOpenPicsForOpt2.Text10.Text = ""
frmOpenPicsForOpt2.Command3.Enabled = False
End Sub

Private Sub Form_Load()
'
Me.Show
'
hl = frmHMHL.Combo2

'
If frmOpenPicsForOpt2.Height = 4320 Then GoTo disp2 'if only one pic was selected....
If frmOpenPicsForOpt2.Height = 4680 Then GoTo disp3 'etc...
If frmOpenPicsForOpt2.Height = 5055 Then GoTo disp4
If frmOpenPicsForOpt2.Height = 5415 Then GoTo disp5
If frmOpenPicsForOpt2.Height = 5745 Then GoTo disp6
If frmOpenPicsForOpt2.Height = 6135 Then GoTo disp7
If frmOpenPicsForOpt2.Height = 6480 Then GoTo disp8
If frmOpenPicsForOpt2.Height = 6825 Then GoTo disp9
If frmOpenPicsForOpt2.Height = 7200 Then GoTo disp10


disp2:
For i = 1 To 1000000 'loop one milliom times - i probably should have use a do...loop but im not good at those so for all my loops i just use for i = 1 to 1000000 'cause usually nobody lets it loop that many times.
On Command1.Value GoTo doneX 'this aint really needed - might as well remove it
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl 'pause for specified num# of seconds that user specified.
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Next i


disp3:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Next i


disp4:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Next i

disp5:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Next i


disp6:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text6.Text)
pause hl
Next i

disp7:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text6.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text7.Text)
pause hl
Next i

disp8:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text6.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text7.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text8.Text)
pause hl
Next i

disp9:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text6.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text7.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text8.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text9.Text)
pause hl
Next i

disp10:
For i = 1 To 1000000
On Command1.Value GoTo doneX
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text1.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text2.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text3.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text4.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text5.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text6.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text7.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text8.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text9.Text)
pause hl
Image1.Picture = LoadPicture(frmOpenPicsForOpt2.Text10.Text)
pause hl
Next i

doneX:


End Sub

Private Sub Image1_Click()
'same as above^
ShowCursor (bShow = False)
'
Me.Hide
frmHMHL.Show
Unload frmOpenPicsForOpt2
'
frmOpenPicsForOpt2.Text1.Text = ""
frmOpenPicsForOpt2.Text2.Text = ""
frmOpenPicsForOpt2.Text3.Text = ""
frmOpenPicsForOpt2.Text4.Text = ""
frmOpenPicsForOpt2.Text5.Text = ""
frmOpenPicsForOpt2.Text6.Text = ""
frmOpenPicsForOpt2.Text7.Text = ""
frmOpenPicsForOpt2.Text8.Text = ""
frmOpenPicsForOpt2.Text9.Text = ""
frmOpenPicsForOpt2.Text10.Text = ""
frmOpenPicsForOpt2.Command3.Enabled = False
End Sub

