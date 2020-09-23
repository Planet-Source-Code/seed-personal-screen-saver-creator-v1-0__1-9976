VERSION 5.00
Begin VB.Form frmOpt4Final 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TEXT WILL GO HERE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   11775
   End
End
Attribute VB_Name = "frmOpt4Final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
frmST.Show
ShowCursor (bShow = False)
End Sub

Private Sub Form_Load()
'loads the specified presets:
Me.BackColor = frmST.CommonDialog2.Color
Label1.ForeColor = frmST.CommonDialog1.Color
Label1.Caption = frmST.Text1.Text
Label1.FontSize = frmST.Text2.Text
ShowCursor (bShow = True)
End Sub

Private Sub Label1_Click()
Unload Me
frmST.Show
ShowCursor (bShow = False)
End Sub
