VERSION 5.00
Begin VB.Form Travel 
   BackColor       =   &H00FF0000&
   Caption         =   "Travel"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFirstLetter 
      Caption         =   "Find Countries that starts with a certain letter"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdTripPrice 
      Caption         =   "Start Planning Your Trip Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmdAlphabet 
      Caption         =   "Sort Countries Alphabetically"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdTravel 
      Caption         =   "Choose The Country You Wish To Travel To See The Airfare Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton cmdLoadandShow 
      Caption         =   "Load and Show Countries"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   4200
      ScaleHeight     =   7635
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "Travel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Travel.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This form lists the countries, sorts them alphabetically
'searches for a the countries that begin with the letter that the user specifies
'the user can click on a country that he or she wishes to travel to and provides the price of a roundtrip ticket (depending on the season)
'The user can start planning for the expenses of his or her trip

Option Explicit
'Dim the variables used by more than one sub command
Dim countries(1 To 30) As String
Dim Airfare(1 To 30) As Single, RunningTotal As Single
Dim CName(1 To 20) As String, NameofCurrency(1 To 20) As String, Curency(1 To 20) As Single
'Hides the Travel Form and Shows the South America Form
Private Sub cmdBack_Click()
Travel.Hide
SouthAmerica.Show
End Sub
'This gives the user the option of searching for a country beginning with a letter of his/her choosing
Private Sub cmdFirstLetter_Click()
Dim Letter As String
Dim X(1 To 100) As String
'Dim the Variables
'Open the data file of countries to put into an array
Open App.Path & "\Travel.txt" For Input As #1
'Using a Do while loop to put the data into an array
ctr = 0
picResults.Cls
Do While Not EOF(1)
            ctr = ctr + 1
                Input #1, countries(ctr), Airfare(ctr)
        Loop
'letter is the specified letter from the user
Letter = InputBox("Type a letter and look for all the countries in South America that start with that letter")
'using a For/Next loop/Exhaustive Search, this sub command searches for as much countries beginning with the
'letter that the user specified and adds them to a counter

For j = 1 To ctr

    If Letter = Left(countries(j), 1) Then
        picResults.Print countries(j)


    End If
    
Next j

End Sub
'Loads the first data file and prints it out in the print box
Private Sub cmdLoadandShow_Click()
cmdTripPrice.Enabled = True
cmdTravel.Enabled = True            'The other buttons are false until this one is clicked, thus making them true
cmdAlphabet.Enabled = True
cmdFirstLetter.Enabled = True

ctr = 0
picResults.Cls  'Clears the printer box

Open App.Path & "\Travel.txt" For Input As #1       'Loads the travel data file into an array with a Do While Loop
    Do While Not EOF(1)
        ctr = ctr + 1
            Input #1, countries(ctr), Airfare(ctr)
    Loop
For j = 1 To ctr
    picResults.Print j, countries(j)        'Prints the countries in the printer box
Next j

Close

End Sub

Private Sub cmdTravel_Click()
Dim country As String, Found As Boolean   'Dim the Variables
i = 0   'Set i to zero
country = InputBox("Which Country Would You Like To Travel To?")    'country prints the desired country as well as the airfare
    
    Do While (Not Found) And i < ctr
        i = i + 1                                       'using a Do While Loop, the specified country is found and stops (match and stop)
            If country = countries(i) Then Found = True
            picResults.Cls
                
    Loop
                
            If Not Found Then
                picResults.Print "Sorry, Your Country Is Not Listed"
            Else
                SelectedCountry = Airfare(i)            'the Airfare price is saved for the expenses in the TripPrice form
                picResults.Print "Country", "Airfare"
                picResults.Print "**************************************"
                picResults.Print countries(i), FormatCurrency(Airfare(i), 0)        'Formats currency ($ in front)
            End If
            
End Sub
'Alphabetizes the Countries by name
Private Sub cmdAlphabet_Click()
Dim Pass As Integer, pos As Integer, TempCountries As String, TempAirfare As Single     'Dim Variables
For Pass = 1 To ctr
    For pos = 1 To ctr - Pass
        If countries(pos) > countries(pos + 1) Then
            TempCountries = countries(pos)          'TempCountries (empty space) makes it possible to move the files into the arranged order
            countries(pos) = countries(pos + 1)
            countries(pos + 1) = TempCountries
        End If
    Next pos
Next Pass
picResults.Cls              'Clear the print box
picResults.Print , "Country"
picResults.Print "*********************************************************************"
For j = 1 To ctr                    'Using an Exhaustive loop, the countries are printed out in alphabetical order
    picResults.Print j, countries(j)
Next j
End Sub
'Hides the Travel Form and Shows the TripPrice Form
Private Sub cmdTripPrice_Click()
TripPrice.Show
Travel.Hide
End Sub
