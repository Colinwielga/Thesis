VERSION 5.00
Begin VB.Form frmboobanna 
   BackColor       =   &H00FFFF00&
   Caption         =   "Anna"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue on your tour de st. joe"
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the Boobery welcome page"
      Height          =   1215
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdchoose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose another person to talk to "
      Height          =   1215
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label lblabout 
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmboobanna.frx":0000
      BeginProperty Font 
         Name            =   "Gill Sans MT Condensed"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label lblanna 
      BackColor       =   &H00FFFF00&
      Caption         =   "Anna"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6420
      Left            =   4920
      Picture         =   "frmboobanna.frx":00FE
      Top             =   480
      Width           =   6570
   End
End
Attribute VB_Name = "frmboobanna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    'Project name:  Tour De St. Joe
    'Form:  frmboobanna, "Anna"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To show who you could be talking to.

Private Sub cmdback_Click()

    frmboob.Show
    frmboobanna.Hide

End Sub

Private Sub cmdchoose_Click()

    frmtalkto.Show
    frmboobanna.Hide
    
End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmboobanna.Hide

End Sub
