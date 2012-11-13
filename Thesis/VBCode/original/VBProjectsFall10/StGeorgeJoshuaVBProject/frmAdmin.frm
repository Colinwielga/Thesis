VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Andminstrator Options"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7905
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00000080&
      Caption         =   "View and Modify Vocabulary Lists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewStudentRecords 
      BackColor       =   &H00000080&
      Caption         =   "View Student Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddNouns 
      BackColor       =   &H00000080&
      Caption         =   "Add Nouns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddVerbs 
      BackColor       =   &H00000080&
      Caption         =   "Add Verbs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblExplaination 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator Options "
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddNouns_Click()
    'Shows the addNouns form
    frmAdmin.Hide
    frmAddNouns.Show
End Sub

Private Sub cmdAddVerbs_Click()
    'shows the add verbs form and hides the admin pane
    frmAddVerbs.Show
    frmAdmin.Hide
End Sub

Private Sub cmdLogOut_Click()
    'Logs the user out using the Log Out Subroutine (see mdlPublicSubs)
    frmAdmin.Hide
    Call LogOut
End Sub

Private Sub cmdModify_Click()
    'shows the classVocab form, and reads the nouns and verbs into arrays for use by frmClassVocab
    frmAdmin.Hide
    frmClassVocab.Show
    'Public Subroutines, cf. mdlPublicSubs
    Call ReadNouns
    Call ReadVerbs
End Sub

Private Sub cmdQuit_Click()
    'Quits the Program
    End
End Sub

Private Sub cmdViewStudentRecords_Click()
    'Shows the student records page
    frmAdmin.Hide
    frmStudentRecords.Show
End Sub
