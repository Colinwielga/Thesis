VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No, I'm broke and can't afford it."
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes!"
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WOULD YOU LIKE TO OPEN YOUR OWN STORE?"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7050
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
    MsgBox ("When you can afford to start your store come back!")
    End
End Sub

Private Sub cmdYes_Click()
    frmStore.Visible = True
    frmWelcome.Visible = False
End Sub
