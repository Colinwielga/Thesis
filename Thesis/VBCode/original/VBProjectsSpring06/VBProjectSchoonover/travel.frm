VERSION 5.00
Begin VB.Form frmMainPage 
   BackColor       =   &H80000013&
   Caption         =   "Schoonover - Main Page"
   ClientHeight    =   7890
   ClientLeft      =   1695
   ClientTop       =   2130
   ClientWidth     =   10095
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Bradley Hand ITC"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10095
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   720
      TabIndex        =   5
      Text            =   "Vacation Planner"
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label lblFootnote 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Please note that vacations depart from the Minneapolis Airport"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   7440
      Width           =   6615
   End
   Begin VB.Label lblHiking 
      BackColor       =   &H80000013&
      Caption         =   "Hiking Vacation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblSki 
      BackColor       =   &H80000013&
      Caption         =   "Ski Vacation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblAdventure 
      BackColor       =   &H80000013&
      Caption         =   "Adventure Vacation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblBeach 
      BackColor       =   &H80000013&
      Caption         =   "Beach Vacation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image imgHiking 
      Height          =   2250
      Left            =   5520
      Picture         =   "travel.frx":0000
      Top             =   4560
      Width           =   2190
   End
   Begin VB.Image imgSki 
      Height          =   1845
      Left            =   1680
      Picture         =   "travel.frx":10212
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Image imgBeach 
      Height          =   2145
      Left            =   1320
      Picture         =   "travel.frx":1A988
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Image imgAdventure 
      Height          =   1995
      Left            =   5640
      Picture         =   "travel.frx":2D9AA
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on the image of the type of vacation you would like to go on"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Vacation Planner(travel.vbp)
'Form Name: frmAdventureVacation (TravelAdventure.frm)
'Author: Nicole Schoonover
'Date: Friday, March 24, 2006
'Objective: The main page allows users to select the type of vaction they wish
    'to travel on (Beach, Adventure, Ski, or Hiking).  Once navigated to these
    'respected pages, users also have the ability to return back to this main page
    'and view any other pages they wish to look at.

Private Sub imgBeach_Click()
'Clicking this image brings users to the Beach Vacation Form
    frmMainPage.Hide
    frmBeachVacation.Show
End Sub

Private Sub imgAdventure_Click()
'Clicking this image brings users to the Adventure Vacation From
    frmMainPage.Hide
    frmAdventureVacation.Show
End Sub

Private Sub imgSki_Click()
'Clicking this image brings users to the Ski Vacation Form
    frmMainPage.Hide
    frmSkiVacation.Show
End Sub

Private Sub imgHiking_Click()
'Clicking this image brings users to the Hiking Vacation Form
    frmMainPage.Hide
    frmHikingVacation.Show
End Sub

