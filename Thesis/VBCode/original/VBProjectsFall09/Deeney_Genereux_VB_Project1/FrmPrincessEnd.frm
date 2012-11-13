VERSION 5.00
Begin VB.Form FrmPrincessEnd 
   BackColor       =   &H00000000&
   Caption         =   "Princess End"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Thanks for helping out the princess!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   600
      Picture         =   "FrmPrincessEnd.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z Z Z... you wore the princess out! Now it is time for her to go to sleep until her next day of shopping fun! "
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   10095
   End
End
Attribute VB_Name = "FrmPrincessEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This ends the princess story



Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdQuit_Click()
    MsgBox "This is where your story successfully ends! Start Over", , "Story Ends"
    FrmPrincessEnd.Hide
    frmWelcome.Show
    
End Sub
