VERSION 5.00
Begin VB.Form frmDestinations 
   BackColor       =   &H000000FF&
   Caption         =   "Destinations"
   ClientHeight    =   12660
   ClientLeft      =   1815
   ClientTop       =   1680
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   ScaleHeight     =   12660
   ScaleWidth      =   14850
   Begin VB.CommandButton cmdReturnToHomePage 
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11160
      TabIndex        =   13
      Top             =   10560
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
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
      Left            =   1680
      TabIndex        =   12
      Top             =   10560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCanada 
      Caption         =   "Canada"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   9
      Top             =   10680
      Width           =   1695
   End
   Begin VB.PictureBox pbxCanada 
      Height          =   1215
      Left            =   8280
      Picture         =   "ParadiseCruisesDestinations.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   10560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMexico 
      Caption         =   "Mexico"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   7
      Top             =   9000
      Width           =   1695
   End
   Begin VB.PictureBox pbxMexico 
      Height          =   1215
      Left            =   8280
      Picture         =   "ParadiseCruisesDestinations.frx":0E15
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdHawaii 
      Caption         =   "Hawaii"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   5
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox pbxHawaii 
      Height          =   1215
      Left            =   8280
      Picture         =   "ParadiseCruisesDestinations.frx":1F3F
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCaribbean 
      Caption         =   "Caribbean"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   3
      Top             =   5640
      Width           =   1695
   End
   Begin VB.PictureBox pbxBahamas 
      Height          =   1215
      Left            =   8280
      Picture         =   "ParadiseCruisesDestinations.frx":3183
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlaska 
      Caption         =   "Alaska"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.PictureBox pbxAlaska 
      Height          =   1215
      Left            =   8280
      Picture         =   "ParadiseCruisesDestinations.frx":4238
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
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
      Height          =   615
      Left            =   12960
      TabIndex        =   14
      Top             =   12000
      Width           =   1575
   End
   Begin VB.Label lblDestinationDirections 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Click on the name of the destination to find out more about it!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5400
      TabIndex        =   11
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label lblDestinations 
      BackColor       =   &H000000FF&
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmDestinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmDestinations (ParadiseCruisesDestinations.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Form: To display pictures of the various destinations
                'and to tell the user more about each place so that
                'they can make a choice about which they want to
                'chose when they are making itinerary choices
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdAlaska_Click()
    'Brings up a message box which gives the user more information about Alaska
    MsgBox "Enjoy stunning vistas of snow-capped mountains & majestic blue-ice glaciers!  Departing from Seward aboard your Southbound Alaska cruise, you’ll see the Great Land in the grandest possible way. This breathtaking and fascinating cruise is more than a seven day vacation; it’s a lifetime of memories! ", , "Alaska"
End Sub
Private Sub cmdCanada_Click()
    'Brings up a message box which gives the user more information about Canada
    MsgBox "Canada’s breathtaking beauty will embrace you.  From the growing cities of Quebec and Toronto to the peaceful mountains of the Canadian Rockies and Stanley Park in Vancouver, Canada’s untainted beauty will astonish you as you near the Arctic Circle and North Pole. The Butchart Gardens of British Columbia will stun you with its natural and well kept fauna.", , "Canada"
End Sub
Private Sub cmdCaribbean_Click()
    'Brings up a message box which gives the user more information about Caribbean
    MsgBox "Explore Mayan Ruins. Relax on a sun-drenched beach or swim with colorful fish.  The exotic Southern Caribbean route visits St. Maarten, Barbados and Martinique; the equally gorgeous tropical destinations in the exotic Western Caribbean itinerary are Belize, Costa Rica and Colon, Panama. Between May and October of 2003, the ship will sail from the heart of Manhattan to the Eastern Caribbean: San Juan, St. Thomas and Tortola/Virgin Gorda.", , "Caribbean"
End Sub
Private Sub cmdHawaii_Click()
    'Brings up a message box which gives the user more information about Hawaii
    MsgBox "Be dazzled by nature as you cruise through this Polynesian paradise.  Sail from Honolulu (O'ahu) and take in the fabulous sights, sounds, fragrances and tastes of exotic ports of call on the islands of Hawai'i, Maui, Kaua'i and Fanning Island in the Republic of Kirabati. Snorkel over a submerged volcano or bicycle down another; dance and feast at a lu'au or bask on a secluded black sand beach. Embrace the spirit of aloha!", , "Hawaii"
End Sub
Private Sub cmdMexico_Click()
    'Brings up a message box which gives the user more information about Mexico
    MsgBox "Magnificent beaches, colorful markets, lively cantinas -- your Mexican Riviera cruise vacation has fun in store for everyone! Adventure abounds in each port of call along the Mexican Riviera. Enjoy sport fishing, mountain biking and snorkeling or sunbathe on one of Mexico's pristine beach resorts and sample the local cuisine.", , "Mexico"
End Sub
Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub
Private Sub cmdReturnToHomePage_Click()
    'Hides the destination form and shows the home form so that the user can chose from
    'other options besides Destination
    frmDestinations.Hide
    frmHome.Show
End Sub
