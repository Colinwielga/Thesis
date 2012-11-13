VERSION 5.00
Begin VB.Form frmboobbrooke 
   BackColor       =   &H00000000&
   Caption         =   "Brooke"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FF80FF&
      Caption         =   "Continue on your tour de st. joe"
      Height          =   1095
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdboobery 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to the Boobery's welcome page"
      Height          =   1095
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdtalk 
      BackColor       =   &H00FF80FF&
      Caption         =   "Choose another person to talk to"
      Height          =   1095
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label lblabout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A little neurotic and self deprecating, but honestly, who wouldn't want to hang out with someone like this... "
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblbrooke 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brooke"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro H"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   6795
      Left            =   2880
      Picture         =   "frmboobbrooke.frx":0000
      Top             =   240
      Width           =   9060
   End
End
Attribute VB_Name = "frmboobbrooke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Project name:  Tour De St. Joe
    'Form:  frmboobbrooke, "Brooke"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To show who you could be talking to.
    
Private Sub cmdboobery_Click()

    frmboob.Show
    frmboobbrooke.Hide

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmboobbrooke.Hide

End Sub

Private Sub cmdtalk_Click()

    frmtalkto.Show
    frmboobbrooke.Hide

End Sub
