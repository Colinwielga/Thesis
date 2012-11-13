VERSION 5.00
Begin VB.Form AndreaFreemanfrmStartUp 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbegin 
      Caption         =   "Let's Begin the Program!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Picture         =   "frmStartUp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Computer Science 130 - Section #2"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmStartUp.frx":2C67
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "March 15, 2004"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Andrea Freeman's Personality Analysis Project "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "AndreaFreemanfrmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjPersonalityAnalysis (Andrea Freeman's VB Project.vbp)
'Form Name: AndreaFreemanfrmStartUp (frmStartUp.frm)
'Author: Andrea Freeman
'Date Written: March 11, 2004
'Purpose of Form: This form states the purpose of the project and provides the user
                  'with an introductory form and a button to proceed to the actual
                  'beginning of the project.

Private Sub cmdbegin_Click()
AndreaFreemanfrmStartUp.Hide 'Hide the first form.
AndreaFreemanfrmFavoriteAnimal.Show 'Show the second form.
End Sub

Private Sub Form_Load()

End Sub
