VERSION 5.00
Begin VB.Form frmSTART 
   BackColor       =   &H00000000&
   Caption         =   "Weezer!"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   Picture         =   "frmSTART.frx":0000
   ScaleHeight     =   7770
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No. . . ."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3840
      TabIndex        =   2
      Top             =   6000
      Width           =   3255
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Heck yes!!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label lblStart 
      BackColor       =   &H00000000&
      Caption         =   "Are you ready to rock out with Weezer?"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   7215
   End
End
Attribute VB_Name = "frmSTART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name; Weezer
'Form Name: frmSTART.frm
'Author: Emily Balamut
'Date Written: 11/4/08
'Objective: This is the beginning of my project. I ask the user if they want
'to start it or not. If so, then they go to the first form (frmBeginning.frm).
'If not, the program ends.
Option Explicit

Private Sub cmdNo_Click()
MsgBox "That's too bad! Come again soon!", , "Sad Day"
End
End Sub

Private Sub cmdYes_Click()
    frmSTART.Hide
    frmBeginning.Show
    
    UserName = InputBox("What's your name?", , "Name?")
    MsgBox "I hope that you have a great time rocking out with Weezer, " & UserName & "!", , "Enjoy!"
    
End Sub
