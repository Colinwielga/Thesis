VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00400000&
      Caption         =   "Are you ready to rock out with Weezer?"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdNo_Click()
MsgBox "That's too bad. Hope you stop by again soon!", , "Bye!"
End
End Sub

Private Sub cmdYes_Click()
    UserName = InputBox("That's AWESOME! Now, what's your name?", "Welcome!")
    frmWelcome.Hide
    frmSTART.Show
End Sub
