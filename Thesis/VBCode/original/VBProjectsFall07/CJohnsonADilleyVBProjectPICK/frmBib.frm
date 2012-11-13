VERSION 5.00
Begin VB.Form frmBib 
   BackColor       =   &H0000FFFF&
   Caption         =   "Work Cited"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C0C0&
      Caption         =   "All Done? Then Click me!"
      Height          =   1095
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Our incredible grasp of Visual Basic"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Label lblRef3 
      BackColor       =   &H0000FFFF&
      Caption         =   "http://www.ctduiattorney.com/dui_information/calculating_bac.html "
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4680
      TabIndex        =   4
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label lblRef2 
      BackColor       =   &H0000FFFF&
      Caption         =   "http://www.ou.edu/oupd/bac.htm "
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label lblRef1 
      BackColor       =   &H0000FFFF&
      Caption         =   "http://www.alcohol.vt.edu/Students/alcoholEffects/estimatingBAC/index.htm "
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblThanks 
      BackColor       =   &H0000FFFF&
      Caption         =   "Thanks for testing your Blood Alcohol Level with ""Drinking Buddy""!"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   3480
      TabIndex        =   1
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Stuff we used to get our information:"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmBib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
'when clicked, the command button ends the program
End
End Sub
