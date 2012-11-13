VERSION 5.00
Begin VB.Form frmboobwhit 
   BackColor       =   &H00008000&
   Caption         =   "whitney"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Continue on your tour de st. joe"
      Height          =   1215
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   2895
   End
   Begin VB.CommandButton cmdboob 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to the Boobery's welcome page"
      Height          =   1215
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   3015
   End
   Begin VB.CommandButton cmdtalk 
      BackColor       =   &H0000FFFF&
      Caption         =   "Choose someone else to talk to"
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Height          =   1215
      Left            =   2880
      TabIndex        =   7
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Height          =   1215
      Left            =   4200
      TabIndex        =   6
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label lblabout 
      BackColor       =   &H000080FF&
      Caption         =   $"frmboobwhit.frx":0000
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Whitney"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblroom 
      BackColor       =   &H00008000&
      Height          =   8055
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   2760
      Picture         =   "frmboobwhit.frx":00CB
      Top             =   -1680
      Width           =   9000
   End
End
Attribute VB_Name = "frmboobwhit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Project name:  Tour De St. Joe
    'Form:  frmboobwhit, "Whitney"
    'Author:  Brooke
    'Date:  3/11/08
    'Objective: To show who you could be talking to.
    
Private Sub cmdboob_Click()

    frmboob.Show
    frmboobwhit.Hide

End Sub

Private Sub cmdleave_Click()

    frmjoetown.Show
    frmboobwhit.Hide

End Sub

Private Sub cmdtalk_Click()

    frmtalkto.Show
    frmboobwhit.Hide

End Sub


