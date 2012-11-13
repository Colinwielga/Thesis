VERSION 5.00
Begin VB.Form DifferentTeams 
   Caption         =   "Form2"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLocate 
      Caption         =   "Locate!"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtLocation 
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Array"
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   4200
      ScaleHeight     =   6315
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblLocation 
      Caption         =   "Where is (team) located?"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
End
Attribute VB_Name = "DifferentTeams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
