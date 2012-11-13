VERSION 5.00
Begin VB.Form frmPopularAttractions 
   Caption         =   "Popular Attractions"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Popular_Attractions_Form.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortAll 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sort All Attractions"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdOther 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdMonarchyGovernment 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Monarchy and Goverment Attractions"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdTheatre 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Theatre"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdMuseums 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Museums"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H00808080&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lblPopularAttractions 
      BackStyle       =   0  'Transparent
      Caption         =   "London's Popular Attractions"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   9375
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Click the button of the types attractions you would like to learn more about."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   5520
      Width           =   9375
   End
End
Attribute VB_Name = "frmPopularAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London
'Form Name: PopularAttractions
'Author: Heather Arnhalt
'Date Written: October 15, 2009
'Objective: Allow the user to select which types of attractions they would like to learn more about.
'Then display options for what they can do or see under that category on a new form for each category

Private Sub cmdGoToHome_Click(Index As Integer)
    'returns the user to the home page of the project
    frmHomePage.Show
    frmPopularAttractions.Hide
End Sub
Private Sub cmdMonarchyGovernment_Click(Index As Integer)
    'Hide the Popular Attractions form and show the Monarchy and Government form
    frmMonarchyGovernment.Show
    frmPopularAttractions.Hide
End Sub

Private Sub cmdMuseums_Click(Index As Integer)
    'Hide the Popular Attractions form and show the Monarchy and Government form
    frmPopularAttractions.Hide
    frmMuseums.Show
End Sub

Private Sub cmdParks_Click(Index As Integer)
    'Hide the Popular Attractions form and show the Parks Form
    frmPopularAttractions.Hide
    frmParks.Show
End Sub

Private Sub cmdOther_Click()
    'Hide the Popular Attractions form and show the Other form
    frmPopularAttractions.Hide
    frmOtherAttractions.Show
End Sub

Private Sub cmdQuit_Click()
    'End the program
    End
End Sub

Private Sub cmdSortAll_Click()
    'hide the Popular Attractions form and show the Sort Attractions Form
    frmPopularAttractions.Hide
    frmSortAttractions.Show
End Sub

Private Sub cmdTheatre_Click(Index As Integer)
    'hide the popular attractions form and show the theatre form
    frmPopularAttractions.Hide
    frmTheatre.Show
End Sub

