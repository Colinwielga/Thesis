VERSION 5.00
Begin VB.Form frmDeathbutheir 
   Caption         =   "Death"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   Picture         =   "frmDeathbutheir.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "End "
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdworkscited 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click to view bibliography"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "End "
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   3015
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmDeathbutheir.frx":1749D
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6720
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmDeathbutheir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form conveys to the user the ultimate outcome of his decisions throughout the game
'via a label and also gives him a quit command button to exit the program

Private Sub Command2_Click()
End
End Sub

