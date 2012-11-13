VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00000000&
   Caption         =   "About"
   ClientHeight    =   8355
   ClientLeft      =   2310
   ClientTop       =   1455
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10845
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00808080&
      Caption         =   "exit"
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H00808080&
      Caption         =   "Return to main"
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox picabout 
      Height          =   6735
      Left            =   0
      Picture         =   "frmabout.frx":0000
      ScaleHeight     =   6675
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click anywhere to learn about our project"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmabout(frmabout.frm)
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'This form is designed to give general information about the project.
Option Explicit


Private Sub cmdexit_Click()
End
End Sub
'Brings you back to Main Screen
Private Sub cmdmain_Click()
    frmabout.Hide
    frmmain.Show
End Sub
'gives info about the project
Private Sub Picabout_Click()
    MsgBox "This program was intended solely for the purpose of entertainment for those who choose to use it.  It was created by Andy Lebovsky and Clay Wilfahrt for a computer science class."
End Sub


