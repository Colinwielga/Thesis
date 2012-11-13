VERSION 5.00
Begin VB.Form frmMsgBoxRussian 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Russian"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Trends"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblRussian1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRussian.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lblRussian 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmMsgBoxRussian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Message Box Russian
'Form Objective: This form appears as a message box with a description of the Russian style when the Russian picture is selected off of the Trends page.

Private Sub cmdReturn_Click()
'This command button allows the user to return to the trends page after viewing the message box with the descripton of the Russian style.
    frmTrends.Show
    frmMsgBoxRussian.Hide
End Sub
