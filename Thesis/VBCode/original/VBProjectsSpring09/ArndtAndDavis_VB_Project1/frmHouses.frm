VERSION 5.00
Begin VB.Form frmHouses 
   BackColor       =   &H00800000&
   Caption         =   "Where Do You See Yourself Living?"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmHouses.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDuplex 
      Caption         =   "View Duplex Price"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12360
      TabIndex        =   16
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox Picture5 
      Height          =   2295
      Left            =   8520
      Picture         =   "frmHouses.frx":06C1
      ScaleHeight     =   2235
      ScaleWidth      =   3195
      TabIndex        =   15
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton cmdRambler 
      Caption         =   "View Rambler Price"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      TabIndex        =   14
      Top             =   1440
      Width           =   2655
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   8400
      Picture         =   "frmHouses.frx":3BBB
      ScaleHeight     =   2355
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdMobileHome 
      Caption         =   "View Mobile Home Price"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   12
      Top             =   8400
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   480
      Picture         =   "frmHouses.frx":7833
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   7800
      Width           =   3495
   End
   Begin VB.CommandButton cmdMansion 
      Caption         =   "View Mansion Price"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   10
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdApartment 
      Caption         =   "View Apartment Price"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   480
      Picture         =   "frmHouses.frx":C799
      ScaleHeight     =   2355
      ScaleWidth      =   3435
      TabIndex        =   7
      Top             =   4800
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   480
      Picture         =   "frmHouses.frx":F70A
      ScaleHeight     =   3315
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Submit House and Continue"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11040
      TabIndex        =   5
      Top             =   9120
      Width           =   4095
   End
   Begin VB.TextBox txtHouse 
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   9480
      Width           =   3855
   End
   Begin VB.CommandButton cmdReturnFinerThings 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   1
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   12720
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Look At Some Fine Real Estate Options For You!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   13215
   End
   Begin VB.Label lblA 
      BackColor       =   &H00800000&
      Caption         =   "Enter Your Ideal House:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   8760
      Width           =   4095
   End
End
Attribute VB_Name = "frmHouses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: The Game Of Life
'Form Name: frmHouses
'Authors: Pam Arndt and Alisa Davis
'Date Written: 3/4/09
'Objective: User views possible house options and their price ranges, then chooses a house that is saved for the summary page
Option Explicit

Private Sub cmdApartment_Click()
MsgBox "Not a bad choice! Rent varies by location, expect to pay more for rent in bigger cities!", , "Apartment"
End Sub

Private Sub cmdDuplex_Click()
MsgBox "Town Homes usually cost between $265,000 and $650,000", , "Town Home"
End Sub

Private Sub cmdEnter_Click()
Dim Choice As Integer

House = txtHouse.Text
Choice = InputBox("Type a 1 if you'd like to find your dream car or 2 if you want to take a chance with the lottery.", "Next")

'An else-if-then statement directs the user to the correct form of their choice.
    If Choice = 1 Then
        frmHouses.Hide
        frmCars.Show
    ElseIf Choice = 2 Then
        frmHouses.Hide
        frmLucky.Show
    Else: MsgBox "Sorry, you entered an invalid option plese enter either a 1 or 2.", , "Error"
End If
End Sub

Private Sub cmdHomepage_Click()
frmHouses.Hide
frmBeginning.Show 'return to beginning form
End Sub

Private Sub cmdMansion_Click()
'Select Case Function to guess a mansion price
Dim BestGuess As Integer
BestGuess = InputBox("Do you know how much the most expensive mansion in the world is worth? (Guess in millions)", "Take a Guess")
    Select Case BestGuess
    Case Is < 75
        MsgBox "Too low!", , "Try Again"
    Case Is > 75
        MsgBox "Not quite that much!", , "Try Again"
    Case Is = 75
        MsgBox "Did you look that up?", , "Correct!"
    Case Else
        MsgBox "Living the lifestyle of the rich and famous with a sweet mansion!", , "Wow!"
    End Select
End Sub

Private Sub cmdMobileHome_Click()
MsgBox "Mobile Homes aren't that bad. They usually cost less than $75,000", , "Trailer Park"
End Sub

Private Sub cmdQuit_Click()
End 'quit program
End Sub

Private Sub cmdRambler_Click()
MsgBox "Way to be all American! Cost is approximately $200,000 to $400,000.", , "Rambler"
End Sub

Private Sub cmdReturnFinerThings_Click()
frmHouses.Hide
frmTheFinerThingsInLife.Show 'back one form
End Sub
