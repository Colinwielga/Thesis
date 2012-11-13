VERSION 5.00
Begin VB.Form frmHomePage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "London Attractions Home Page"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Home_Page_Form.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPopularAttractions 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Popular Attractions"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdTheatre 
      BackColor       =   &H00C0FFFF&
      Caption         =   "London Theatre"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   4920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdCurrency 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Convert Currency"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLondonAttractions 
      BackStyle       =   0  'Transparent
      Caption         =   "London Attractions"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   840
      TabIndex        =   4
      Top             =   6840
      Width           =   5895
   End
End
Attribute VB_Name = "frmHomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London
'Form Name: Home Page
'Author: Heather Arnhalt
'Date Written: October 15, 2009
'Form Objective: To allow the user to decide which aspect of London they would like
'to learn more about. The user selects what they would like to learn more about and
'the program takes them to a page that will provide them with information
'Project Objective: The objective of this project is to provide the user with various
'tidbits about London attractions and how to get there.

Private Sub cmdCurrency_Click(Index As Integer)
    'show the Currency form and hide the Home Page form
    frmCurrency.Show
    frmHomePage.Hide
End Sub

Private Sub cmdPopularAttractions_Click(Index As Integer)
    'show the popular attractions form and hide the other forms
    frmHomePage.Hide
    frmTheatre.Hide
    frmPopularAttractions.Show
End Sub

Private Sub cmdQuit_Click(Index As Integer)
    'ends the program
    End
End Sub

Private Sub cmdTheatre_Click(Index As Integer)
    'show the theatre form and hide the other forms
    frmHomePage.Hide
    frmPopularAttractions.Hide
    frmTheatre.Show
End Sub

