VERSION 5.00
Begin VB.Form Stratus 
   BackColor       =   &H80000008&
   Caption         =   "Dodge Stratus"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   120
      Picture         =   "frm11.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Dodge Stratus"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   2895
   End
End
Attribute VB_Name = "Stratus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Stratus.Hide
End Sub

