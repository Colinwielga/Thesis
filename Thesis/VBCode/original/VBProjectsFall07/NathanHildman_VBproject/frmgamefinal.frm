VERSION 5.00
Begin VB.Form frmgamefinal 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Final Case"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6240
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdplayerscase 
      Caption         =   "YOUR CASE!!"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Text            =   "Only Two Cases left!!  Make your Final Decision"
      Top             =   600
      Width           =   9975
   End
End
Attribute VB_Name = "frmgamefinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdplayerscase_Click()
frmgamefinal.Hide                   'goes to another form
frmplayerscase.Show



End Sub

Private Sub Command1_Click()
frmgamefinal.Hide                   'goes to another form
frmgamefinalcase.Show


End Sub
