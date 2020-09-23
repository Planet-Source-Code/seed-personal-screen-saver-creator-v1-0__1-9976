VERSION 5.00
Begin VB.Form frmOpt1Final 
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
Attribute VB_Name = "frmOpt1Final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Click()
frmOpenPicForOpt1.Show
Me.Hide
ShowCursor (bShow = False) 'show mouse
End Sub

Private Sub Image1_Click()
frmOpenPicForOpt1.Show
Me.Hide
ShowCursor (bShow = False) 'show mouse
End Sub

