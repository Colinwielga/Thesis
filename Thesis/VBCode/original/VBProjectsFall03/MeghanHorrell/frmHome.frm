VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00C0C000&
   Caption         =   "Home"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   16
      Top             =   10440
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10920
      Width           =   2655
   End
   Begin VB.PictureBox pbxEntertainment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      Picture         =   "frmHome.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   8400
      Width           =   1335
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
      Left            =   9720
      Picture         =   "frmHome.frx":1534
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   8640
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
      Picture         =   "frmHome.frx":2237
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   11
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
      Picture         =   "frmHome.frx":37B4
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   10
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
      Picture         =   "frmHome.frx":4878
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   9
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
      Picture         =   "frmHome.frx":5474
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8640
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdLifeOnBoard 
      BackColor       =   &H00FF8080&
      Caption         =   "Life on Board"
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
      TabIndex        =   4
      Top             =   8640
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
      Picture         =   "frmHome.frx":6529
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
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
      TabIndex        =   14
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
Option Explicit
Private Sub cmdAboutCriseLine_Click()
    frmAboutParadiseCruises.Show
    frmHome.Hide
End Sub

Private Sub cmdCruiseOptions_Click()
    frmCruiseOptions.Show
    frmHome.Hide
End Sub

Private Sub cmdDestinations_Click()
    frmDestinations.Show
    frmHome.Hide
End Sub

Private Sub cmdItineraryOptions_Click()
frmHome.Hide
frmItinerary.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRooms_Click()
    frmRooms.Show
    frmHome.Hide
End Sub
