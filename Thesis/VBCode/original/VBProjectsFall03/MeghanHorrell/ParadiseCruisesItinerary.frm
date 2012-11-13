VERSION 5.00
Begin VB.Form frmItinerary 
   BackColor       =   &H0000C000&
   Caption         =   "Display of Itinerary Options"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   1680
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdGoToCruiseOptionsPage 
      BackColor       =   &H00FF00FF&
      Caption         =   "Go to Cruise Options Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9360
      Width           =   1695
   End
   Begin VB.PictureBox pbxResultsAllOptions 
      BeginProperty Font 
         Name            =   "NIST Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   12840
      ScaleHeight     =   8115
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   3840
      Width           =   6135
   End
   Begin VB.TextBox txtDays 
      Height          =   975
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturnToHome 
      BackColor       =   &H00FF00FF&
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
      Height          =   1095
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10920
      Width           =   2295
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10920
      Width           =   2175
   End
   Begin VB.TextBox txtDestination 
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtPrice 
      Height          =   975
      Left            =   3600
      TabIndex        =   2
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FF00FF&
      Caption         =   "Display your Itinterary Options"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2175
   End
   Begin VB.PictureBox pbxResultsItinerary 
      BeginProperty Font 
         Name            =   "NIST Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   6240
      ScaleHeight     =   8115
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H0000C000&
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
      Left            =   240
      TabIndex        =   18
      Top             =   13200
      Width           =   2895
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   $"ParadiseCruisesItinerary.frx":0000
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   9360
      Width           =   4215
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "To enter more than one set of choices, simply delete your previous choices and type in your new ones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   7440
      TabIndex        =   15
      Top             =   8400
      Width           =   3975
   End
   Begin VB.Label lblAllOptions 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Here are all of the options you can pick from according to the ones you have chosen"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   14
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label lblCurrentSelection 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Your Current Selection Options Are"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   13
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label lblDays 
      BackColor       =   &H0000C000&
      Caption         =   "Enter the number of days you would like to travel for (either 7, 5 or 4)"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblItineraryOptions 
      BackColor       =   &H0000C000&
      Caption         =   "Itinerary Options"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label lblDestination 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Enter a destination that you are looking for"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Enter a Price that you are looking for(without decimals or commas:  i.e.""1700"")"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Click this button to display your itinerary options"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmItinerary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmItinerary (ParadiseCruisesItinerary.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Form: 'This form allows the user to enter the criteria they are looking
                'for when they are making their travel plans with the number of days
                'they are looking to travel for, the price they are looking for and
                'the destination they are looking for and then by clicking on the
                'Display your itinerary" button, the user is able to see which options
                'are available to them according to what the cruise line has to offer.
                'On this form the user is able to see the results of their current
                'search in one box while at the same time, they can see the results
                'of all of their searches in the other box
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Public PATH As String
'Declares the Variables for the path and deminsions them
Private Sub cmdDisplay_Click()
'Declares the variables locally
Dim Days As Integer
Dim PriceEntered As Single
Dim DestinationEntered As String
'Dimensions the arrays for CruiseDestination (type String),Suite(type String),
'and Price(type Integer) from 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim Flag As Boolean
'Initializes the flag as false so that when the user inputs something in the textbox
'it will only print when the criteria are met, meaning that the flag is then set equal
'to True
Flag = False
'Declares the variable "Days" as whatever the user inputs into the text box "txtDays.Text"
Days = txtDays.Text
'Sets the condition that the data file will only be opened if the user enters a 7
If Days = 7 Then
    'Opens the data file "7 day Cruises.txt"
    Open PATH & "7 day Cruises.txt" For Input As #1
        'Reads the data file into an array
        For i = 1 To 30
            Input #1, CruiseDestination(i), Suite(i), Price(i)
        Next i
    'Closes the data file
    Close #1
    'Declares the variable "PriceEntered" as whatever the user inputs into the text box
    '"textPrice.Text"
    PriceEntered = txtPrice.Text
    'Declares the variable "DestinationEntered" as whatever the user inputs into the
    'text box "txtDestination.Text"
    DestinationEntered = txtDestination.Text
    'Prints a line above each heading so that it is easier to read
    pbxResultsAllOptions.Print "----------------------------------------------------------------------------------------------"
    'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
    pbxResultsAllOptions.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
    'Prints a line below each heading so that it is easier to read
    pbxResultsAllOptions.Print "----------------------------------------------------------------------------------------------"
        For i = 1 To 30
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 30) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itinerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
'Sets the condition that the data file will only be opened if the user enters a 5
ElseIf Days = 5 Then
        'Opens the data file "5 day Cruises.txt"
        Open PATH & "5 day Cruises.txt" For Input As #1
            'Reads the data file into an array
            For i = 1 To 30
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        'Closes the data file
        Close #1
        'Declares the variable "PriceEntered" as whatever the user inputs into the text box
        '"textPrice.Text"
        PriceEntered = txtPrice.Text
        'Declares the variable "DestinationEntered" as whatever the user inputs into the
        'text box "txtDestination.Text"
        DestinationEntered = txtDestination.Text
        'Prints a line above each heading so that it is easier to read
        pbxResultsAllOptions.Print "-------------------------------------------------------------------------------------------"
        'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
        pbxResultsAllOptions.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
        'Prints a line below each heading so that it is easier to read
        pbxResultsAllOptions.Print "-------------------------------------------------------------------------------------------"
        For i = 1 To 30
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 30) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
'Sets the condition that the data file will only be opened if the user enters a 4
ElseIf Days = 4 Then
        'Opens the data file "4 day Cruises.txt"
        Open PATH & "4 day Cruises.txt" For Input As #1
            'Reads the data file into an array
            For i = 1 To 18
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        'Closes the data file
        Close #1
        'Declares the variable "PriceEntered" as whatever the user inputs into the text box
        '"textPrice.Text"
        PriceEntered = txtPrice.Text
        'Declares the variable "DestinationEntered" as whatever the user inputs into the
        'text box "txtDestination.Text"
        DestinationEntered = txtDestination.Text
        'Prints a line above each heading so that it is easier to read
        pbxResultsAllOptions.Print "--------------------------------------------------------------------------------------------"
        'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
        pbxResultsAllOptions.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
        'Prints a line below each heading so that it is easier to read
        pbxResultsAllOptions.Print "--------------------------------------------------------------------------------------------"
        For i = 1 To 18
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 18) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsAllOptions.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
End If
'Clears the itinerary box so that with each new set of options the user entered, they
'are able to just see thier current selection rather then all of the selections at once
pbxResultsItinerary.Cls
Days = txtDays.Text
'Sets the condition that the data file will only be opened if the user enters a 7
If Days = 7 Then
    'Opens the data file "7 day Cruises.txt"
    Open PATH & "7 day Cruises.txt" For Input As #1
        'Reads the data file into an array
        For i = 1 To 30
            Input #1, CruiseDestination(i), Suite(i), Price(i)
        Next i
    'Closes the data file
    Close #1
    'Declares the variable "PriceEntered" as whatever the user inputs into the text box
    '"textPrice.Text"
    PriceEntered = txtPrice.Text
    'Declares the variable "DestinationEntered" as whatever the user inputs into the
    'text box "txtDestination.Text"
    DestinationEntered = txtDestination.Text
    'Prints a line above each heading so that it is easier to read
    pbxResultsItinerary.Print "----------------------------------------------------------------------------------------------"
    'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
    pbxResultsItinerary.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
    'Prints a line below each heading so that it is easier to read
    pbxResultsItinerary.Print "----------------------------------------------------------------------------------------------"
        For i = 1 To 30
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 30) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itinerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
'Sets the condition that the data file will only be opened if the user enters a 5
ElseIf Days = 5 Then
        'Opens the data file "5 day Cruises.txt"
        Open PATH & "5 day Cruises.txt" For Input As #1
            'Reads the data file into an array
            For i = 1 To 30
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        'Closes the data file
        Close #1
        'Declares the variable "PriceEntered" as whatever the user inputs into the text box
        '"textPrice.Text"
        PriceEntered = txtPrice.Text
        'Declares the variable "DestinationEntered" as whatever the user inputs into the
        'text box "txtDestination.Text"
        DestinationEntered = txtDestination.Text
        'Prints a line above each heading so that it is easier to read
        pbxResultsItinerary.Print "-------------------------------------------------------------------------------------------"
        'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
        pbxResultsItinerary.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
        'Prints a line below each heading so that it is easier to read
        pbxResultsItinerary.Print "-------------------------------------------------------------------------------------------"
        For i = 1 To 30
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 30) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
'Sets the condition that the data file will only be opened if the user enters a 4
ElseIf Days = 4 Then
        'Opens the data file "5 day Cruises.txt"
        Open PATH & "4 day Cruises.txt" For Input As #1
            'Reads the data file into an array
            For i = 1 To 18
                Input #1, CruiseDestination(i), Suite(i), Price(i)
            Next i
        'Closes the data file
        Close #1
        'Declares the variable "PriceEntered" as whatever the user inputs into the text box
        '"textPrice.Text"
        PriceEntered = txtPrice.Text
        'Declares the variable "DestinationEntered" as whatever the user inputs into the
        'text box "txtDestination.Text"
        DestinationEntered = txtDestination.Text
        'Prints a line above each heading so that it is easier to read
        pbxResultsItinerary.Print "--------------------------------------------------------------------------------------------"
        'Prints the headings "Cruise Destination","Suite" and "Price of Trip"
        pbxResultsItinerary.Print "Cruise Destination"; Tab(25); "Suite"; Tab(45); "Price of Trip"
        'Prints a line below each heading so that it is easier to read
        pbxResultsItinerary.Print "--------------------------------------------------------------------------------------------"
        For i = 1 To 18
            'Sets the conditions that if the price the user enters is greater than or
            'equal to the price in the data file and the destination the user enters
            'equals the destination in the data file, then the flag is equal to true,
            'meaning that the if both the price and the destination of a particular line
            'in the data file (1 to 18) fit the criteria then it will print the line in
            'the results box and then move to the next line
            If PriceEntered >= Price(i) And DestinationEntered = CruiseDestination(i) Then
                Flag = True
                pbxResultsItinerary.Print CruiseDestination(i); Tab(25); Suite(i); Tab(45); FormatCurrency(Price(i))
            End If
        Next i
        'If none or one of these criteria are met for any of the lines of the data file,
        'then the flag is false and a message box pops up letting the user know that
        'what they entered does not match up with any of the options
        If Flag = False Then
            MsgBox "Sorry, we do not have an itenerary that fits your price and destination selection.  If you are still interested, please select another price or go to our Cruise Options page to see our options.", , "Sorry"
        End If
    
End If
End Sub

Private Sub cmdGoToCruiseOptionsPage_Click()
    'Hides the itinerary form and shows the Cruise Options form if the user needs to refer back to it
    frmItinerary.Hide
    frmCruiseOptions.Show
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdReturnToHome_Click()
    'Hides the itinerary form and shows the Home Page
    frmItinerary.Hide
    frmHome.Show
End Sub

Private Sub Form_Load()
    PATH = "N:\CS130\handin\Meghan Horrell\"
End Sub
