VERSION 5.00
Begin VB.Form BMW 
   BackColor       =   &H80000008&
   Caption         =   "BMW 530i"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "frm27.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "BMW 530i"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "BMW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    BMW.Hide
End Sub

