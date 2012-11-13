VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalc 
      Caption         =   "4: Calculate Price"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "3: Input Specifications"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "2: Show Other Options"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdRailings 
      Caption         =   "1: Show Railings"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.PictureBox picDisplay 
      Height          =   3615
      Left            =   3000
      ScaleHeight     =   3555
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "David Schonhardt"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AutoBid (\Schonhardt_David\Bid.vbp)
'MainForm (Bid.frm)
'Written by David Schonhardt on October 17-21, 2003 for the general pricing of a deck to be added to the side of a house based on a limited set of variables for the general reference of a homeowner.
'Form1 is for the general program, the form within which all the inputs are recieved and calculations are done.
'Declares all the variables used by multiple buttons
Dim SurfaceType As Integer, RailingType As Integer
Dim Angles As Integer, Stairs As Integer, LandWidth As Single, LandLength As Single
Dim DeckLength As Single, DeckWidth As Single, DeckHeight As Single

Private Sub cmdRailings_Click()
'Brings up the window with the railing types.
Railings.Show
MainForm.Hide
End Sub

Private Sub cmdOptions_Click()
'Brings up window with the various deck options.
Options.Show
MainForm.Hide
End Sub

Private Sub cmdInput_Click()
DeckLength = InputBox("Enter the Length of the Deck") 'Inputs the length of the deck in feet.
DeckWidth = InputBox("Enter the Width of the Deck") 'Inputs the width of the deck in feet.
DeckHeight = InputBox("Enter the Height of the Deck") 'Inputs the height of the deck in feet.
SurfaceType = InputBox("Enter the Type of Decking Preferred (1 for Trex, 2 for TimberTech, 3 for Cedar)") 'Selects the decking material.
RailingType = InputBox("Enter the Type of Railing Preferred (1 for Post, 2 for Flat, 3 for Privacy Wall, 4 for Vinyl)") 'Selects the railing material.
Angles = InputBox("Enter the Number of Angled Corners") 'Selects the number of angled corners.
Stairs = InputBox("Enter 1 for a Staircase, 0 for no Staircase") 'Selects whether or no the user wants a staircase.
LandWidth = InputBox("Enter the Width of the Landing (If no landing is planned, then enter 0)") 'Inputs the width of the lower stair landing, if one is planned.
LandLength = InputBox("Enter the Length of the Landing (If no landing is planned, then enter 0)") 'Inputs the length of the lower stair landing, if one is planned.
End Sub

Private Sub cmdCalc_Click()
picDisplay.Cls
'Calculates the cost of the platform (Height, Length, Width, Material and Angles).
Open Path & "decking.txt" For Input As #1
Dim Decking(1 To 4)
Dim PlatformCost As Double, CTR As Integer
CTR = 1
Do While Not EOF(1)
    Input #1, Decking(SurfaceType)
    CTR = CTR + 1
Loop
PlatformCost = (DeckLength * DeckWidth) * Decking(SurfaceType) + (20 * Angles) + (3 * DeckHeight)
Select Case SurfaceType
    Case Is = 1
        picDisplay.Print "The Cost of a Trex Platform is"; Tab(31); FormatCurrency(PlatformCost)
    Case Is = 2
        picDisplay.Print "The Cost of a TimberTech Platform is"; Tab(31); FormatCurrency(PlatformCost)
    Case Is = 3
        picDisplay.Print "The Cost of a Cedar Platform is"; Tab(31); FormatCurrency(PlatformCost)
End Select
Close #1
'Calculates the cost of the railing (2xLength of deck, Width, Material of Railing).
Open Path & "railings.txt" For Input As #2
Dim Railing(1 To 5)
Dim RailingCost As Double
CTR = 1
Do While Not EOF(2)
    Input #2, Railing(RailingType)
    CTR = CTR + 1
Loop
RailingCost = (2 * DeckLength + DeckWidth) * Railing(RailingType)
Select Case RailingType
    Case Is = 1
        picDisplay.Print "The Cost of a Post Railing is"; Tab(31); FormatCurrency(RailingCost)
    Case Is = 2
        picDisplay.Print "The Cost of a Flat Railing is"; Tab(31); FormatCurrency(RailingCost)
    Case Is = 3
        picDisplay.Print "The Cost of a Privacy Wall is"; Tab(31); FormatCurrency(RailingCost)
    Case Is = 4
        picDisplay.Print "The Cost of a Vinyl Railing is"; Tab(31); FormatCurrency(RailingCost)
End Select
Close #2
'Calculates the cost of stairs (Height, Materials, Railing Material).
Dim StairCost As Double
If Stairs = 1 Then
    StairCost = DeckHeight * 30
End If
If StairCost > 0 Then
    picDisplay.Print "The Cost of a Staircase is"; Tab(31); FormatCurrency(StairCost)
End If
'Calculates the cost of a landing (Landing Width, Landing Length, Material).
Dim LandingCost As Double
If LandWidth > 0 Then
    LandingCost = (LandWidth * LandLength) * (Decking(SurfaceType) * 1.5)
End If
If LandingCost > 0 Then
    picDisplay.Print "The Cost of a Landing is"; Tab(31); FormatCurrency(LandingCost)
End If
'Calculates the total cost.
picDisplay.Print "-------------------------------------------------------------------"
Dim DeckCost As Double
DeckCost = LandingCost + StairCost + RailingCost + PlatformCost
DeckCost = Round(DeckCost)
picDisplay.Print "The Cost of the Deck is"; Tab(31); FormatCurrency(DeckCost)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

