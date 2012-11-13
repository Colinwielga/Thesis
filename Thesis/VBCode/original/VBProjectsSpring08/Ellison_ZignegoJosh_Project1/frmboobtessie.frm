VERSION 5.00
Begin VB.Form frmboobtessie 
   BackColor       =   &H00C000C0&
   Caption         =   "Tessie"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Continue on your tour de st. joe"
      Height          =   1575
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton cmdboob 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Return to the Boobery home page"
      Height          =   1575
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdtalk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Choose someone else to talk to"
      Height          =   1575
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmboobtessie.frx":0000
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lbltessie 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tessie"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6795
      Left            =   4680
      Picture         =   "frmboobtessie.frx":00AE
      Top             =   -120
      Width           =   9060
   End
End
Attribute VB_Name = "frmboobtessie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Project name:  Tour De St. Joe
    'Form:  frmboobtessie, "Tessie"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To show who you could be talking to.

Private Sub cmdboob_Click()

    frmboob.Show
    frmboobtessie.Hide

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmboobtessie.Hide

End Sub

Private Sub cmdtalk_Click()

    frmtalkto.Show
    frmboobtessie.Hide

End Sub
