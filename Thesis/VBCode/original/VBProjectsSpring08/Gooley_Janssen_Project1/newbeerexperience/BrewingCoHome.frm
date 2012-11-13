VERSION 5.00
Begin VB.Form Companies 
   BackColor       =   &H0000C000&
   Caption         =   "Brewing Companies Home Page"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form2"
   ScaleHeight     =   7260
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00C000C0&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdMiller 
      BackColor       =   &H000000FF&
      Caption         =   "Miller Brewing Co."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton cmdCoors 
      BackColor       =   &H00FF0000&
      Caption         =   "Coors Brewing Co."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton cmdAB 
      BackColor       =   &H0000FFFF&
      Caption         =   "Anheuser-Busch Co."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Companies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Beer Experience
'Companies
'Lauren Gooley and Tim Janssen
'March 21, 2008
'This form is used to navigate between the various companies that we researched for our project.
'Each subroutine causes a specific form to be hidden or to be shown

Private Sub cmdAB_Click()
frmAB.Show
Companies.Hide
End Sub

Private Sub cmdCoors_Click()
frmCoors.Show
Companies.Hide
End Sub

Private Sub cmdMainMenu_Click()
Companies.Hide
frmStartUp.Show
End Sub

Private Sub cmdMiller_Click()
frmMiller.Show
Companies.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub
