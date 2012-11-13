VERSION 5.00
Begin VB.Form Asscroftloses 
   BackColor       =   &H004D3CC1&
   Caption         =   "How did this happen??"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton QuitAshcroftL 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Programmed by Megan Kelly"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H004D3CC1&
      Caption         =   $"Asscroftloses.frx":0000
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   3885
      Left            =   1080
      Picture         =   "Asscroftloses.frx":00ED
      Top             =   240
      Width           =   4905
   End
End
Attribute VB_Name = "Asscroftloses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Quit_Click()
End
End Sub
