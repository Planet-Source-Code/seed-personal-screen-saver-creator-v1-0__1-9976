VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Screen Saver Wizard "
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "mail me with any questions or comments:  Aedseed@aol.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SeeD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuPicBasedScrSaver 
            Caption         =   "Picture-Based Screen Saver"
         End
         Begin VB.Menu mnuTxtBasedScrSaver 
            Caption         =   "Text-Based Screen Saver"
         End
      End
      Begin VB.Menu mnuBltInScrSavers 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuFuture 
         Caption         =   "Future?"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_DblClick()
frm2.Show
Me.Hide
End Sub

Private Sub mnuBltInScrSavers_Click()
frmSBI.Show
Me.Hide
End Sub

Private Sub mnuExit_Click()
'unload all the forms!
Unload Me
Unload frmSelectPicBasedScrSaver
Unload frmOpenPicsForOpt2
Unload frmOpenPicForOpt1
Unload frmHMHL
Unload frmOpt1Final
Unload frmOpt2Final
Unload frmSBI
Unload frmopt3final
Unload frmOpt4Final
Unload frmDir
Unload frmOpenPicForOpt3
Unload frmST

End Sub

Private Sub mnuFuture_Click()
MsgBox "i hope you like this program.  it was not hard to make but there was a lot of coding to it and i  made it all myself from scratch.  i use vb6 LE and win95 and this program works good on my computer.  if you like this program or have any questions or comments then mail me at aedseed@aol.com.  If this program gets good response, i'll add to it.  i hope to add things like password protection and other stuff too.  if you have any ideas then email me at aedseed@aol.com.  please vote for this at planet source code if you like it!  i might add more options and user defined features to this.  thanks!    -SeeD", vbOKOnly, "Future:"
End Sub

Private Sub mnuPicBasedScrSaver_Click()
frmSelectPicBasedScrSaver.Show
Me.Hide
End Sub

Private Sub mnuStTxt_Click()
frmST.Show
Me.Hide
End Sub

Private Sub mnuTxtBasedScrSaver_Click()
frmST.Show
End Sub
