VERSION 5.00
Begin VB.Form frmSkiVacation 
   BackColor       =   &H00800000&
   Caption         =   "Schoonover - Ski Vacation"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Height          =   615
      Left            =   3840
      TabIndex        =   14
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Calculate the total cost of airfare and hotel"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter the number of people traveling and the number of nights you wish to stay"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton cmdResults 
      BackColor       =   &H008080FF&
      Caption         =   "Show All Vacations In This Price Range"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   2895
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   9195
      TabIndex        =   9
      Top             =   3960
      Width           =   9255
   End
   Begin VB.CommandButton cmdSpecials 
      BackColor       =   &H0000FF00&
      Caption         =   "See Ski Vacation Airfare Specials!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   2895
   End
   Begin VB.OptionButton optLots 
      BackColor       =   &H00FFFF80&
      Caption         =   "More than $1,000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton optHigh 
      BackColor       =   &H00FFFF80&
      Caption         =   "$750.01 - $1,000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton optMedHigh 
      BackColor       =   &H00FFFF80&
      Caption         =   "$500.01 - $750"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optMed 
      BackColor       =   &H00FFFF80&
      Caption         =   "$250.01 - $500"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optLow 
      BackColor       =   &H00FFFF80&
      Caption         =   "$0 - $250"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblInput 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Type the location of the vacation that you would like to go on"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00FFFF80&
      Caption         =   "What is the maximum amount you are willing to spend per person?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label lblSki 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ski Vacation"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSkiVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Vacation Planner(travel.vbp)
'Form Name: frmSkiVacation (TravelAdventure.frm)
'Author: Nicole Schoonover
'Date: Friday, March 24, 2006
'Objective:
    'This form allows users to identify the maximum amount of money they
    'are willing to spend per person on a ski vacation.  Also, they can enter
    'the number of people traveling and the number of nights they wish to stay.
    'Once this information is entered, they can click a button that will
    'calculate the price per person (adding the airfare plus the nightly cost
    'per person) and display it in the output box.  After the information is
    'displayed, the user can then type in the location they wish to travel to
    '(word for word) and the program will calculate how much that trip will
    'cost (grand total).  If the user enters an invalid location an "error"
    'message will be displayed.  In addition to displaying this information,
    'users can see the weekly specials.

Dim People As Integer
Dim Nights As Integer

Private Sub cmdCalculate_Click()
'Calculate the total cost of the vacation

Dim Choice As String
Dim Pos, Size, GrandTotal, Temp As Integer
Dim Location(1 To 100) As String
Dim Airfare(1 To 100) As Single
Dim Hotel(1 To 100) As Single
Choice = txtInput
Temp = 0

picOutput.Cls

'Read the input file
Open App.Path & "\SkiVacations.txt" For Input As #1
Pos = 0
Do Until EOF(1)
    Pos = Pos + 1
    Input #1, Location(Pos), Airfare(Pos), Hotel(Pos)
Loop
Close #1

Size = Pos

picOutput.Print "Number of people traveling: " & People,
picOutput.Print
picOutput.Print "Number of nights in hotel: " & Nights,
picOutput.Print

'Find a string match for the users desired vacation location
For Pos = 1 To Size
    If Choice = Location(Pos) Then
        GrandTotal = (Airfare(Pos) * People) + (Hotel(Pos) * Nights)
        Temp = Temp + 1
        picOutput.Print "The total cost for airfare and hotel is: " & FormatCurrency(GrandTotal)
        Pos = Size
    End If
Next Pos

'Display message box if there are no exact matches
If Temp = 0 Then
    MsgBox "Entry not found.  Make sure you type the location EXACTLY as it was on the screen.", , "ERROR!"
End If
End Sub

Private Sub cmdInput_Click()
'Input the number of people and the number of nights

    People = InputBox("Enter the number of people traveling", "Number of People")
    Nights = InputBox("Enter the number of nights you would like to stay", "Number of Nights")
End Sub

Private Sub cmdResults_Click()
'Display the results of vacations that are less than the max. price range

Dim Pos, Size, Temp As Integer
Dim Location(1 To 100) As String
Dim Airfare(1 To 100) As Single
Dim Hotel(1 To 100) As Single
Dim HotelPersonNightly, TotalPersonCost As Single

picOutput.Cls

'Input File
Open App.Path & "\SkiVacations.txt" For Input As #1
Pos = 0
Do Until EOF(1)
    Pos = Pos + 1
    Input #1, Location(Pos), Airfare(Pos), Hotel(Pos)
Loop
Close #1

Size = Pos

picOutput.Print "LOCATION"; Tab(40); "AIRFARE"; Tab(58); "HOTEL"; Tab(70); "TOTAL COST PER PERSON"

'Read through the file and calculate using the input values for number of people
'and number of nights and do the corresponding action to that value.
For Pos = 1 To Size
    HotelPersonNightly = ((Hotel(Pos) * Nights) / People)
    TotalPersonCost = Airfare(Pos) + HotelPersonNightly
        If optLots = True And TotalPersonCost > 1000 Then
            picOutput.Print Location(Pos); Tab(40); FormatCurrency(Airfare(Pos)); Tab(58); FormatCurrency(Hotel(Pos)); Tab(70); FormatCurrency(TotalPersonCost)
            Temp = Temp + 1
        ElseIf optHigh = True And TotalPersonCost <= 1000 Then
            picOutput.Print Location(Pos); Tab(40); FormatCurrency(Airfare(Pos)); Tab(58); FormatCurrency(Hotel(Pos)); Tab(70); FormatCurrency(TotalPersonCost)
            Temp = Temp + 1
        ElseIf optMedHigh = True And TotalPersonCost <= 750 Then
            picOutput.Print Location(Pos); Tab(40); FormatCurrency(Airfare(Pos)); Tab(58); FormatCurrency(Hotel(Pos)); Tab(70); FormatCurrency(TotalPersonCost)
            Temp = Temp + 1
        ElseIf optMed = True And TotalPersonCost <= 500 Then
            picOutput.Print Location(Pos); Tab(40); FormatCurrency(Airfare(Pos)); Tab(58); FormatCurrency(Hotel(Pos)); Tab(70); FormatCurrency(TotalPersonCost)
            Temp = Temp + 1
        ElseIf optLow = True And TotalPersonCost <= 250 Then
            picOutput.Print Location(Pos); Tab(40); FormatCurrency(Airfare(Pos)); Tab(58); FormatCurrency(Hotel(Pos)); Tab(70); FormatCurrency(TotalPersonCost)
            Temp = Temp + 1
        End If
Next Pos

'If there are no matches display the following message box
If Temp = 0 Then
    MsgBox "There are no matches to your search criteria.", , "ERROR!"
End If


End Sub

Private Sub cmdReturn_Click()
'Go back to the main page

    frmSkiVacation.Hide
    frmMainPage.Show
End Sub

Private Sub cmdSpecials_Click()
'Display the weekly specials by reading the appropriate input file

Dim Location(1 To 100) As String
Dim Price(1 To 100) As Single
Dim Pos As Integer
picOutput.Cls
picOutput.Print "Ski Vacation Specials For This Week:"
Open App.Path & "\WeeklySkiSpecials.txt" For Input As #1
Pos = 0
Do Until EOF(1)
    Pos = Pos + 1
    Input #1, Location(Pos), Price(Pos)
    picOutput.Print Location(Pos); Tab(40); FormatCurrency(Price(Pos))
Loop
Close #1
End Sub
