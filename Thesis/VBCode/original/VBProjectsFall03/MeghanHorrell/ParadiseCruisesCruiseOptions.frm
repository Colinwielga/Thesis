VERSION 5.00
Begin VB.Form frmCruiseOptions 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Cruise Options"
   ClientHeight    =   11010
   ClientLeft      =   555
   ClientTop       =   285
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSuiteFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   13320
      Width           =   2175
   End
   Begin VB.CommandButton cmdDestinationFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   11880
      Width           =   2415
   End
   Begin VB.CommandButton cmdSuiteFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton cmdDestinationFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdSuiteSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by Suite!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdDestinationSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by destination!"
      BeginProperty Font 
         Name            =   "Gungsuh"
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
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturnToHomePage 
      BackColor       =   &H00FF00FF&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   12120
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   12120
      Width           =   2415
   End
   Begin VB.PictureBox pbxResults 
      BeginProperty Font 
         Name            =   "NIST Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   7200
      ScaleHeight     =   9675
      ScaleWidth      =   9915
      TabIndex        =   7
      Top             =   2160
      Width           =   9975
   End
   Begin VB.CommandButton cmdPriceFour 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 4 day cruises sorted by price!"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11880
      Width           =   2535
   End
   Begin VB.CommandButton cmdPriceFive 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 5 day cruises sorted by price"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton cmdPriceSeven 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to see our selection of 7 day cruises sorted by price!"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H00FFC0C0&
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
      Height          =   735
      Index           =   1
      Left            =   16200
      TabIndex        =   17
      Top             =   12360
      Width           =   1575
   End
   Begin VB.Label lblMyName 
      BackColor       =   &H00FF0000&
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
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblFourDay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4 day cruises"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   5
      Top             =   10680
      Width           =   5535
   End
   Begin VB.Label lblFiveDay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5 day cruises"
      BeginProperty Font 
         Name            =   "Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   6600
      Width           =   5535
   End
   Begin VB.Label lblSevenDays 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7 day cruises"
      BeginProperty Font 
         Name            =   "@Gungsuh"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lblCruiseOptions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cruise Options"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   12015
   End
End
Attribute VB_Name = "frmCruiseOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjParadiseCruises (Meghan Horrell's VB Project.vbp)
'Form Name : frmCruiseOptions (ParadiseCruisesCruiseOptions.frm)
'Author: Meghan Horrell
'Date Written For: October 29, 2003
'Purpose of Form: Displays the user's cruise options for 7 day, 5 day and 4 day
                'cruises according to Cruise Destination,Suite and Price
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
'Declares the variables for the path and dimensions them
Public PATH As String
Private Sub cmdDestinationFive_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "5 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "5 day Cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
'Closes the data file
Close #1
    'Sorts the data according to Cruise Destination in alphabetical order
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the Cruise Destination in alphabetical order and prints the corresponding suite and price
        'from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdDestinationFour_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 18
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 18
N = 18
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "4 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "4 day cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Cruise Destination in alphabetical order
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        'Prints the Cruise Destination in alphabetical order and prints the corresponding suite and price
        'from the first entry to the 18th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdDestinationSeven_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "7 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "7 day cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Cruise Destination in alphabetical order
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If CruiseDestination(i) > CruiseDestination(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the Cruise Destination in alphabetical order and prints the corresponding suite and price
        'from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub
Private Sub cmdPriceFive_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "5 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "5 day cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Price in from least to greatest
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the from the least to the greatest and prints the corresponding cruise destination and suite
        'from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdPriceFour_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 18
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 18
N = 18
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "4 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "4 day cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Price from least to greatest
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        'Prints the Price in from least to greatest and prints the corresponding suite and price
        'from the first entry to the 18th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdReturnToHomePage_Click()
    'Shows the destinations form and hides the home form
    frmDestinations.Hide
    frmHome.Show
End Sub
Private Sub cmdPriceSeven_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "7 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "7 day cruises.txt" For Input As #1
    'Reads the Data file into an array
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Price from least to greatest
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Price(i) > Price(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the Price in from least to greatest and prints the corresponding destination and price
        'from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteFive_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "5 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "5 day cruises.txt" For Input As #1
    'Sorts the data according to Suite in alphabetical order and from most expensive to least expensive suite
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Suite in alphabetical order and from most expensive to least expensive
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the Suite in alphabetical order and from most expensive to least expensive
        'and prints the corresponding destination and price from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteFour_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 18
Dim CruiseDestination(1 To 18) As String
Dim Suite(1 To 18) As String
Dim Price(1 To 18) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 18
N = 18
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "4 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "4 day cruises.txt" For Input As #1
    'Sorts the data according to Suite in alphabetical order and from most expensive to least expensive suite
    For i = 1 To 18
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Suite in alphabetical order and from most expensive to least expensive
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 18
        'Prints the Suite in alphabetical order and from most expensive to least expensive
        'and prints the corresponding destination and price from the first entry to the 18th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub cmdSuiteSeven_Click()
'Dimensions the arrays locally for CruiseDestination (type String), Suite (type String) and Price(type Integer) for 1 to 30
Dim CruiseDestination(1 To 30) As String
Dim Suite(1 To 30) As String
Dim Price(1 To 30) As Integer
'Declares the variables locally
Dim i As Integer
Dim pass As Integer
Dim tempPrice As Integer
Dim tempCruiseDestination As String
Dim tempSuite As String
Dim N As Integer
'Initializes N as 30
N = 30
'Clears the results of the picture box every time a new button is pushed
pbxResults.Cls
'prints the headings "Cruise Destination","Suite" and "Price" at the top of the picture box
pbxResults.Print "Cruise Destination"; Tab(40); "Suite"; Tab(70); "Price"
'Prints a line under the headings to make them more visible
pbxResults.Print "---------------------------------------------------------------------------------------------------------------------------------"
'Opens the data file "7 day Cruises.txt" and for the arrays that are used in the form CruiseOptions
Open PATH & "7 day cruises.txt" For Input As #1
    For i = 1 To 30
        Input #1, CruiseDestination(i), Suite(i), Price(i)
    Next i
    'Closes the data file
    Close #1
    'Sorts the data according to Suite in alphabetical order and from most expensive to least expensive
    For pass = 1 To N - 1
        For i = 1 To N - pass
            If Suite(i) > Suite(i + 1) Then
                tempPrice = Price(i)
                tempCruiseDestination = CruiseDestination(i)
                tempSuite = Suite(i)
                Price(i) = Price(i + 1)
                CruiseDestination(i) = CruiseDestination(i + 1)
                Suite(i) = Suite(i + 1)
                Price(i + 1) = tempPrice
                CruiseDestination(i + 1) = tempCruiseDestination
                Suite(i + 1) = tempSuite
            End If
        Next i
    Next pass
    For i = 1 To 30
        'Prints the Suite in alphabetical order and from most expensive to least expensive
        'and prints the corresponding destination and price from the first entry to the 30th entry
        pbxResults.Print CruiseDestination(i); Tab(40); Suite(i); Tab(70); FormatCurrency(Price(i))
    Next i
End Sub

Private Sub Form_Load()
    'Shows how to access the path for data files
    PATH = "N:\CS130\handin\Meghan Horrell\"
End Sub
