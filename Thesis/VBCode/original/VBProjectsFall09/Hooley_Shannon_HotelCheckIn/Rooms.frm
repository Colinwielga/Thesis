VERSION 5.00
Begin VB.Form frmRooms 
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   Picture         =   "Rooms.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoToCheckIn 
      BackColor       =   &H0080FFFF&
      Caption         =   "Return To Check In"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   6360
      ScaleHeight     =   6555
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdPresidential 
      BackColor       =   &H0080FFFF&
      Caption         =   "Presidential Suite"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdSuite 
      BackColor       =   &H0080FFFF&
      Caption         =   "Suite"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdKing 
      BackColor       =   &H0080FFFF&
      Caption         =   "King Bed"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdQueen 
      BackColor       =   &H0080FFFF&
      Caption         =   "Queen Bed"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdDouble 
      BackColor       =   &H0080FFFF&
      Caption         =   "Double Bed"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Hotel Check In
'frmRooms
'Shannon Hooley
'10/16/09
'this forms gives an in depth description on pricing and what is included for the different types of rooms

Private Sub cmdDouble_Click()
'clears the previous room recap
picResults.Cls
'prints the info for double rooms
picResults.Print "Double Bed Rooms come standard with:"
picResults.Print "**********************"
picResults.Print "Two Double Beds"
picResults.Print "A desk with in room telephone"
picResults.Print "A flat screen TV with cable"
picResults.Print "Internet hookups"
picResults.Print "Bathroom with complementary towels"
picResults.Print "An ironing board and iron"
picResults.Print "A safe to store your valuables"
picResults.Print "      $109 a night"
End Sub

Private Sub cmdGoToCheckIn_Click()
'brings the guest back to the check in
frmRooms.Hide
frmCheckIn.Show
End Sub

Private Sub cmdKing_Click()
'clears the info from previous room recap
picResults.Cls
'prints info for king rooms
picResults.Print "King Bed Rooms come standard with:"
picResults.Print "**********************"
picResults.Print "One King sized bed"
picResults.Print "A desk with in room telephone"
picResults.Print "A flat screen TV with cable"
picResults.Print "Internet hookups"
picResults.Print "Bathroom with complementary towels"
picResults.Print "An ironing board and iron"
picResults.Print "A safe to store your valuables"
picResults.Print "      $206 a night"
End Sub

Private Sub cmdPresidential_Click()
'clears info from previous recap
picResults.Cls
'prints info for pres. suites
picResults.Print "Presidential Suites come standard with:"
picResults.Print "**********************"
picResults.Print "Two King sized beds"
picResults.Print "A desk with in room telephone"
picResults.Print "A flat screen TV with cable"
picResults.Print "Internet hookups"
picResults.Print "Bathroom with complementary towels"
picResults.Print "An ironing board and iron"
picResults.Print "A safe to store your valuables"
picResults.Print "A living room with a couch and love seat"
picResults.Print "A flat screen TV in both the living room and bedroom"
picResults.Print "Complementary room service"
picResults.Print "A full kitchen (utensils included)"
picResults.Print "Wet bar"
picResults.Print "Mini fridge stocked with:"
picResults.Print "  - a wide array of liquor"
picResults.Print "  - bar snacks"
picResults.Print "  (ALL COMPLIMENTARY)"
picResults.Print "      $453 a night"
End Sub

Private Sub cmdQueen_Click()
'clears info from previous recap
picResults.Cls
'prints info for queen rooms
picResults.Print "Queen Bed Rooms come standard with:"
picResults.Print "**********************"
picResults.Print "One Queen sized bed"
picResults.Print "A desk with in room telephone"
picResults.Print "A flat screen TV with cable"
picResults.Print "Internet hookups"
picResults.Print "Bathroom with complementary towels"
picResults.Print "An ironing board and iron"
picResults.Print "A safe to store your valuables"
picResults.Print "      $150 a night"
End Sub

Private Sub cmdSuite_Click()
'clears info from prev. recap
picResults.Cls
'prints info for suites
picResults.Print "Suites come standard with:"
picResults.Print "**********************"
picResults.Print "Two Queen sized beds"
picResults.Print "A desk with in room telephone"
picResults.Print "A flat screen TV with cable"
picResults.Print "Internet hookups"
picResults.Print "Bathroom with complementary towels"
picResults.Print "An ironing board and iron"
picResults.Print "A safe to store your valuables"
picResults.Print "A living room with a couch and love seat"
picResults.Print "A flat screen TV in both the living room and bedroom"
picResults.Print "      $379 a night"
End Sub
