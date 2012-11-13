VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00C0C000&
   Caption         =   "Home"
   ClientHeight    =   12570
   ClientLeft      =   2040
   ClientTop       =   780
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12570
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdItineraryOptions 
      BackColor       =   &H00FFC0FF&
      Caption         =   "To choose options for your itinerary, click here!"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8400
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10800
      Width           =   2655
   End
   Begin VB.PictureBox pbxRooms 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "ParadiseCruisesHome.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
   End
   Begin VB.PictureBox pbxAboutParadise 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      Picture         =   "ParadiseCruisesHome.frx":0D03
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox pbxDining 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      Picture         =   "ParadiseCruisesHome.frx":2280
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox pbxSpecialOffers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "ParadiseCruisesHome.frx":3344
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.PictureBox pbxDestinations 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "ParadiseCruisesHome.frx":3F40
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdRooms 
      BackColor       =   &H00FF8080&
      Caption         =   "Room Options"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdAboutCriseLine 
      BackColor       =   &H00FF8080&
      Caption         =   "About Paradise Cruises"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdDining 
      BackColor       =   &H00FF8080&
      Caption         =   "Dining"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdCruiseOptions 
      BackColor       =   &H00FF8080&
      Caption         =   "Cruise Options"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton cmdDestinations 
      BackColor       =   &H00FF8080&
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2535
   End
   Begin VB.PictureBox pbxCruiseShipPicture 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4800
      Picture         =   "ParadiseCruisesHome.frx":4FF5
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by Meghan Horrell"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   15
      Top             =   11880
      Width           =   2895
   End
   Begin VB.Label lblChoose 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "To learn more about any of these exciting aspects of our Cruises, just click next to the picture!"
      BeginProperty Font 
         Name            =   "@MS Mincho"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label lblParadiseCruises 
      BackColor       =   &H00C0C000&
      Caption         =   "Paradise Cruises"
      BeginProperty Font 
         Name            =   "Colonna MT"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   10095
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmHome (ParadiseCruisesHome.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Project:  The purpose of this project is to inform the user about Paradise
                    'Cruises.  It gives the user information about the various destinations
                    'they can travel to, about thier dining options, about thier room options,
                    'and about thier travel options according to price, destination and suite.
                    'Then the project allows the user to find out what itinerary options they
                    'have when they input information about the number of days they would like
                    'to travel,the price they are looking for and the destination they are looking
                    'for.

'Purpose of Form: To Display the Options that the user can chose
                'from to find out more about the cruise and to
                'provide the user with  buttons that they can click on,
                'leading them to the forms that give additional information
                'about the cruise.  This form also gives the user a
                'button to click on which allows them to find out various
                'options for their itinerary
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdAboutCriseLine_Click()
    'Shows the About Paradise Cruises form and hides the Home form
    frmAboutParadiseCruises.Show
    frmHome.Hide
End Sub

Private Sub cmdCruiseOptions_Click()
    'Shows the Cruise Options form and hides the Home form
    frmCruiseOptions.Show
    frmHome.Hide
End Sub

Private Sub cmdDestinations_Click()
    'Shows the Destinations form and hides the Home form
    frmDestinations.Show
    frmHome.Hide
End Sub

Private Sub cmdDining_Click()
    'Shows the Dining form and hides the Home form
    frmDining.Show
    frmHome.Hide
End Sub

Private Sub cmdItineraryOptions_Click()
    'Shows the Itinerary Options form and hides the Home form
    frmItinerary.Show
    frmHome.Hide
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdRooms_Click()
    'Shows the Room Options form and hides the Home form
    frmRooms.Show
    frmHome.Hide
End Sub


