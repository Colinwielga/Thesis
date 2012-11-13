VERSION 5.00
Begin VB.Form frmVenues 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H0000FF00&
      Caption         =   "Credits"
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF80&
      Height          =   7695
      Left            =   3120
      ScaleHeight     =   7635
      ScaleWidth      =   11835
      TabIndex        =   6
      Top             =   1800
      Width           =   11895
   End
   Begin VB.CommandButton cmdSBH 
      Caption         =   "Stephen B. Humphrey"
      Height          =   2535
      Left            =   120
      Picture         =   "frmBlueprints.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdColman 
      Caption         =   "Colman"
      Height          =   2535
      Left            =   120
      Picture         =   "frmBlueprints.frx":3DBF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton cmdGorecki 
      Caption         =   "Gorecki"
      Height          =   2535
      Left            =   120
      Picture         =   "frmBlueprints.frx":4F1F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdPetters 
      Caption         =   "Petters"
      Height          =   2535
      Left            =   120
      Picture         =   "frmBlueprints.frx":607F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmBlueprints.frx":71DF
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmVenues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmVenues
'Author: Kurt Oostra
'Date Written:3/27/08
'Objective: Give more information on the CSB/SJU Venues and show the credits
Option Explicit
Private Sub cmdColman_Click()
'clears old information and prints information about the Black Box
picResults.Cls
picResults.Print "The Colman Black Box is the newest venue on campus."
picResults.Print
picResults.Print "The Black Box has a very unique style and has already been used for many different types of shows."
picResults.Print
picResults.Print "Seating in the Black Box varies depending on how many chairs can be safely added after the stage is built for each show."
picResults.Print
End Sub

Private Sub cmdCredits_Click()
'prints the credits for the project
picResults.Cls
picResults.Print "Pictures and Venue Information from www.csbsju.edu"
picResults.Print "Fixture Pictures and Information from"
picResults.Print "www.ETCconnect.com"
picResults.Print "www.strandlighting.com"
picResults.Print "www.altmanltg.com "
picResults.Print "Also for more information on Slide Shows and Check boxes"
picResults.Print "The VB example Multi_form_Sample_w_pictures"
picResults.Print "Nicholas Swanson's VBProject FoodFight 2007"
End Sub

Private Sub cmdGorecki_Click()
'clears old information and prints information about the Gorecki
picResults.Cls
picResults.Print "The Gorecki Family Theater is one of the main theater venues on campus."
picResults.Print
picResults.Print "The Gorecki Family Theater seats 300 patrons."
picResults.Print
picResults.Print "Blueprints and ground plans can be found on the school's website."
End Sub

Private Sub cmdPetters_Click()
'clears old information and prints information about Petters
picResults.Cls
picResults.Print "Petters Auditorium is the largest venue on either campus."
picResults.Print
picResults.Print "It is host to most of the traveling groups that come to perform at CSB/SJU"
picResults.Print
picResults.Print "Petters Auditorium is the only other site that the MN symphony orchestra regularly plays in MN other then their home hall."
picResults.Print
picResults.Print "Petters Auditorium seats 1078 patrons."
picResults.Print
picResults.Print "Blueprints and ground plans can be found on the school's website."
End Sub

Private Sub cmdReturn_Click()
'Returns to Main Menu
frmMainMenu.Show
frmVenues.Hide
End Sub

Private Sub cmdSBH_Click()
'clears old information and prints information about the SBH
picResults.Cls
picResults.Print "The Stephen B. Humphrey is more commonly called the SBH."
picResults.Print
picResults.Print "It is a smaller venue, but is quite capable of being used for almost any type of performance."
picResults.Print
picResults.Print "The SBH seats 515 patrons."
picResults.Print
picResults.Print "Blueprints and ground plans can be found on the school's website."
End Sub
