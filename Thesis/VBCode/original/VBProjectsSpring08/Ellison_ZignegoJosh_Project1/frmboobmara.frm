VERSION 5.00
Begin VB.Form frmboobmara 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Mara"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Continue on your tour de st. joe"
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   2775
   End
   Begin VB.CommandButton cmdboob 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Return to the Boobery home page"
      Height          =   1215
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   2775
   End
   Begin VB.CommandButton cmdtalk 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Choose another person to talk to"
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   3480
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblabout 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmboobmara.frx":0000
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblmara 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mara"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   9060
      Left            =   3840
      Picture         =   "frmboobmara.frx":00B8
      Top             =   -1200
      Width           =   6720
   End
End
Attribute VB_Name = "frmboobmara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Project name:  Tour De St. Joe
    'Form:  frmboobmara, "Mara"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To show who you could be talking to.
    
Private Sub cmdboob_Click()

    frm.boob.Show
    frmboobmara.Hide

End Sub

Private Sub cmdleave_Click()

    frm.joetown.Show
    frm.boobmara.Hide

End Sub

Private Sub cmdtalk_Click()

    frmtalkto.Show
    frmboobmara.Hide

End Sub

