VERSION 5.00
Begin VB.Form frmRooms 
   BackColor       =   &H00C000C0&
   Caption         =   "Rooms"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrandSuite 
      Height          =   3495
      Left            =   1200
      Picture         =   "frmRooms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton cmdRoyalSuite 
      Height          =   3495
      Left            =   1200
      Picture         =   "frmRooms.frx":5086
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   4575
   End
   Begin VB.CommandButton cmdOwnersSuite 
      Height          =   3495
      Left            =   1200
      Picture         =   "frmRooms.frx":AC77
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   4575
   End
   Begin VB.CommandButton cmdSilverSuite 
      Height          =   3495
      Left            =   6720
      Picture         =   "frmRooms.frx":10EEB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton cmdVerandaSuite 
      Height          =   3495
      Left            =   6720
      Picture         =   "frmRooms.frx":1642E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   4575
   End
   Begin VB.CommandButton cmdVistaSuite 
      Height          =   3495
      Left            =   6720
      Picture         =   "frmRooms.frx":1C4A8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   4575
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
      Left            =   12240
      TabIndex        =   2
      Top             =   7800
      Width           =   2415
   End
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
      Left            =   12240
      TabIndex        =   1
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label lblRoomPrices 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      Caption         =   $"frmRooms.frx":217EC
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGrandSuite_Click()
     MsgBox "The Grand Suite!  Like a posh penthouse, the Grand Suite will delight your senses with its whirlpool tub and Bang & Olufsen entertainment center. You'll also enjoy the forward facing windows and verandas that give rise to the most commanding views of the horizon. A majestic window on the world usually reserved for the ship's Captain, now available exclusively to you!", , "Grand Suite"
End Sub
Private Sub cmdOwnersSuite_Click()
    MsgBox "The Owner's Suite!  Step into the Owner's Suite and you will feel as though you have entered a stylish apartment along the shores of the Italian Riviera. Decorated in hushed pastels and warm hues - all of which were personally selected by the ship's owner - this exclusive 827 square-foot hideaway whispers of cozy sophistication.", , "Owner's Suite"
End Sub
Private Sub cmdQuit_Click()
    End
End Sub
Private Sub cmdReturnToHomePage_Click()
    frmRooms.Hide
    frmHome.Show
End Sub

Private Sub cmdRoyalSuite_Click()
    MsgBox "The Royal Suite!  Spacious and spectacular, with 1031 square feet, the Royal Suites live up to their names in dimensions and appointments. With separate living and dining rooms, the Royal Suite is the perfect cloister for cocktail parties, wine tastings or private dinner parties - all of which your suite stewardess will gladly arrange.", , "Royal Suite"
End Sub

Private Sub cmdSilverSuite_Click()
    MsgBox "The Silver Suite!  There are sanctuaries at sea that captivate the senses, that tempt you to linger in their tranquillity and leave you feeling quite peaceful. You will find such places across the threshold of the Silver Suite."
End Sub

Private Sub cmdVerandaSuite_Click()
    MsgBox "The Veranda Suite!  The Veranda Suite includes a teak veranda with floor-to-ceiling glass doors and patio furniture. This suite is beautifully appointed with all the comforts of home, including a marbled bath with a full-sized tub. A walk-in closet with a private safe. A cocktail cabinet that is continuously stocked with your preferences. And feather-down pillows fluffed by your diligent suite stewardess each evening.", , "Veranda Suite"
End Sub

Private Sub cmdVistaSuite_Click()
    MsgBox "The Vista Suite!  The Vista Suite includes a large picture window providing panoramic ocean views. This suite is beautifully appointed with all the comforts of home, including a marbled bath with a full-sized tub. A walk-in closet with a private safe. A cocktail cabinet that is continuously stocked with your preferences. And feather-down pillows fluffed by your diligent suite stewardess each evening.", , "Vista Suite"
End Sub
