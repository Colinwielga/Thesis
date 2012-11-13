VERSION 5.00
Begin VB.Form frmpolicerun 
   BackColor       =   &H000000FF&
   Caption         =   "Game Over"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to exit"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lbldone 
      BackColor       =   &H000000FF&
      Caption         =   "You tried to run away from the police!  GAME OVER."
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmpolicerun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
    'Project name:  Tour De St. Joe
    'Form:  frmpolicerun, "Run"
    'Author:  Brooke
    'Date:  3/30/08
    'Objective: To show what happens when you fuck with the police.

Private Sub cmddone_Click()
    End
End Sub
