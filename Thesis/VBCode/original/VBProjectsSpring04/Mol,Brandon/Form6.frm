VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   Caption         =   "VRSC"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form6"
   ScaleHeight     =   8640
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to return to the main menu"
      Height          =   1575
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   6120
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   480
      Picture         =   "Form6.frx":5E25
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "VRSCB V-Rod"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "VRSCA V-Rod"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   4920
      Width           =   3255
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form6.Hide
End Sub
