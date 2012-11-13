VERSION 5.00
Begin VB.Form frmFlightInformation2 
   BackColor       =   &H00FF00FF&
   Caption         =   "Alaskan Flight Information"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetId 
      BackColor       =   &H0080FF80&
      Caption         =   "Get Id Number"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   6855
      Left            =   3120
      ScaleHeight     =   6795
      ScaleWidth      =   7395
      TabIndex        =   6
      Top             =   840
      Width           =   7455
   End
   Begin VB.CommandButton cmdReturn33 
      BackColor       =   &H00FF80FF&
      Caption         =   "Return to Alaskan Home Page"
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdBook 
      BackColor       =   &H0080FF80&
      Caption         =   "Book your flight"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdCheapest 
      BackColor       =   &H0080FF80&
      Caption         =   "List cheapest flights (less than $250)"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrices 
      BackColor       =   &H0080FF80&
      Caption         =   "Display the flight and price from the city you will be departing from"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdDepartures 
      BackColor       =   &H0080FF80&
      Caption         =   "Display departures and arrivals"
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblFl 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Flight Information"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmFlightInformation2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sunshine & Snow Cruise Lines
'Form Name: frmFlightInformation2
'Authors: Brittany Nosal & Kelly Sunder
'Date Written: 3/15/2009
'Objective: This form includes a command button that displays the major cities in each state from which the user would
'depart from in order to arrive in Miami, Florida to board to cruise ship. There is a command button which enables an
'input box to pop up and asks the user to enter the city and state from which they will be departing. Once this information
'is entered into the inputbox, a messagebox pops up with the price per person for that specific flight. Another command button
'displays the cheapest flights available and their corresponding cities and states (cheapest as in less than $250). The last
'command button books the user's flight for the trip if they decide to do so. An input box appears asking the user to enter
'their desired departure location, and then another inputbox appears asking the user for the total amount of people traveling
'on that particular flight. It then calculates the total for the amount of people given and displays that information in the
'picture box on the right.

Option Explicit
Dim Flights(1 To 100) As String, Arrival(1 To 100) As String, Prices(1 To 100) As Single, CTR As Integer

Private Sub cmdBook_Click()
Dim NumberOfPpl As Integer, runningtotal As Single, J As Integer
Dim x As String
Dim Found As Boolean


x = InputBox("Enter the city and state from which you will be departing from.")
NumberOfPpl = InputBox("Enter the number of people you wish to book the flight for.")
J = 0
Found = False

picResults.Cls

Do While ((Not Found) And (J < CTR))
    J = J + 1
    If x = Flights(J) Then
            Found = True
    End If
Loop

If (Not Found) Then
    MsgBox "That city does not exist on the list or you have the incorrect form. Please enter a city and state from the list in the form City, State."
End If

For J = 1 To CTR
    If x = Flights(J) Then
        runningtotal = NumberOfPpl * Prices(J)
    End If
Next J

picResults.Print "The total cost for"; NumberOfPpl; "person(s) to depart from "; x; " is "; FormatCurrency(runningtotal, 2)
cmdGetId.Enabled = True
End Sub

Private Sub cmdCheapest_Click()
Dim J As Integer, Biggest As Integer

picResults.Cls

picResults.Print "Flights"; Tab(30); "Destination"; Tab(60); "Cost"
picResults.Print "************************************************************************************"

For J = 1 To CTR
    If Prices(J) < 250 Then
        picResults.Print Flights(J); Tab(30); Arrival(J); Tab(60); FormatCurrency(Prices(J), 2)
    End If
Next J


End Sub

Private Sub cmdDepartures_Click()
picResults.Cls

picResults.Print "Flights"; Tab(30); "Destination"; Tab(60)
picResults.Print "*************************************************************"

Open App.Path & "\States2.txt" For Input As #1

CTR = 0

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Flights(CTR), Arrival(CTR), Prices(CTR)
    picResults.Print Flights(CTR); Tab(30); Arrival(CTR)
Loop


Close #1
cmdDepartures.Enabled = False
cmdBook.Enabled = True
cmdPrices.Enabled = True
cmdCheapest.Enabled = True
End Sub

Private Sub cmdGetId_Click()
Dim x As String, id As String, n As Integer
Dim first As String, middle As String, last As String

x = InputBox("Please enter your first name, middle initial, and last name using commas to seperate.")

n = InStr(x, " ")
first = Left(x, n - 1)
last = Right(x, Len(x) - (n + 2))
middle = Mid(x, n + 1, 1)
id = Left(first, 1) & middle & Left(last, 6)

MsgBox "Your id for the cruise ship is " & id
End Sub

Private Sub cmdPrices_Click()
Dim Found As Boolean
Dim City As String
Dim J As Integer

J = 0
Found = False


City = InputBox("Please enter the major city from the list that you will be departing from in the form city, state.")

Do While ((Not Found) And (J < CTR))
    J = J + 1
    If City = Flights(J) Then
        Found = True
        MsgBox "A flight from " & Flights(J) & " to " & Arrival(J) & " will cost " & FormatCurrency(Prices(J), 2) & " per person."
    End If
Loop

If (Not Found) Then
    InputBox ("Please enter a city from the list in the form of city, state.")
End If

End Sub

Private Sub cmdReturn33_Click()
frmAlaskanHome.Show
frmFlightInformation2.Hide
End Sub

