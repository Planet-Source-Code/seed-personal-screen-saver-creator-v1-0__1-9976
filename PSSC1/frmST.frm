VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stationary Text"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "frmST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options:"
      Enabled         =   0   'False
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5055
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "36"
         Top             =   2040
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   720
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Caption         =   "Font Size:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblBC 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Background Color:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblFC 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Font Color:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "What to say:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowColor
lblFC.BackColor = CommonDialog1.Color
End Sub

Private Sub Command2_Click()
CommonDialog2.ShowColor
lblBC.BackColor = CommonDialog2.Color
End Sub

Private Sub Command3_Click()
frmOpt4Final.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Text1_Change()
'make sure that something will be said:
If Text1.Text <> "" Then
 Frame1.Enabled = True
 Label2.Enabled = True
 Label4.Enabled = True
 Label5.Enabled = True
 lblFC.Enabled = True
 lblBC.Enabled = True
 Command1.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = True
 Text2.Enabled = True
 Else
 Frame1.Enabled = False
 Label2.Enabled = False
 Label4.Enabled = False
 Label5.Enabled = False
 lblFC.Enabled = False
 lblBC.Enabled = False
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Text2.Enabled = False
End If
End Sub
