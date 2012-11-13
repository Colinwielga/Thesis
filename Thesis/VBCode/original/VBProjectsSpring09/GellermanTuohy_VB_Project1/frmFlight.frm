VERSION 5.00
Begin VB.Form frmFlight 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16590
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   16590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFlight 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdMatchAndStopSearch 
      Caption         =   "Name the Flight You Want And We Will Tell You How Much It Costs!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtTickets 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextActivityPage 
      BackColor       =   &H000080FF&
      Caption         =   "Let's See What Awesome Activites We Have in Store For You at Your Travel Destination!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3840
      TabIndex        =   4
      Top             =   7080
      Width           =   9855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Awesome Johnnie Travel Experience!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton cmdFlightOptions 
      Caption         =   "View Available Flights"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFF00&
      Height          =   4335
      Left            =   7560
      ScaleHeight     =   4275
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Please enter flight ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "No. of tickets:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Gotta Catch A Flight!!!!!!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   9495
   End
End
Attribute VB_Name = "frmFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/21/09
'This form loads the data from a file into three arrays that will hold the flight ID number, flight takeoff to destination, and the cost for a
'single ticket. There is a match and stop search that is based on the Location that was saved earlier in the program. It won't allow the user to
'book a flight to a location different from the destination that they originally signed up for. Once they are done booking the flight, they will
'then be brought to the activity form that is associated with the location that was saved earlier in the program.

Option Explicit
Dim Tickets As Integer      'declare all the global variables
Dim FlightNum As Integer
Dim GoNext As Boolean

Private Sub cmdFlightOptions_Click()
    Open App.Path & "\Flights.txt" For Input As #1
    
    CTR = 0             'opens the file and loads it into three arrays
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, FlightID(CTR), TakeOff(CTR), TicketCost(CTR)
    Loop
    Close #1            'closes file once it has been loaded into the arrays
    
    'prints the headings before the For/Next loop prints all the data from the file
    picResults.Print "Flight ID Number"; Tab(22); "Flight Departure to Destination"; Tab(55); "Ticket Cost Per Person"
    picResults.Print "************************************************************************************************************"
    
    'For/Next loop prints the flight Id number, where they takeoff and land, and the cost per ticket for all the data from the file
    For I = 1 To CTR
        picResults.Print FlightID(I); Tab(22); TakeOff(I); Tab(55); FormatCurrency(TicketCost(I))
    Next I
    
End Sub

Private Sub cmdMatchAndStopSearch_Click()
    
    Found = False
    I = 0
    
    Tickets = txtTickets.Text
    FlightNum = CInt(txtFlight.Text)
    
    'this part checks for st joe
    
    If Location = "St. Joe" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 1 And FlightNum <= 4 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
            FlightTotal = Tickets * TicketCost(I)
            MsgBox "The total cost for all " & Tickets & " tickets that you have purchased for your flight from " & TakeOff(I) & " is " & FormatCurrency(FlightTotal)
        End If
    End If
    
    'this part checks for Normandy
    
    If Location = "Normandy" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 5 And FlightNum <= 8 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
            FlightTotal = Tickets * TicketCost(I)
            MsgBox "The total cost for all " & Tickets & " tickets that you have purchased for your flight from " & TakeOff(I) & " is " & FormatCurrency(FlightTotal)
        End If
    End If
    
    'this part checks for canada
    
    If Location = "Saskatchewan" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 9 And FlightNum <= 12 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
            FlightTotal = Tickets * TicketCost(I)
            MsgBox "The total cost for all " & Tickets & " tickets that you have purchased for your flight from " & TakeOff(I) & " is " & FormatCurrency(FlightTotal)
        End If
    End If
    
    'this part checks for south dakota
    
    If Location = "Badlands" Then
        Do While ((Not Found) And (I < CTR))    'searches for flight ID number and then once it has found an ID number,
            I = I + 1                           'it saves the flight ID number as I and Found becomes True
            If FlightNum = FlightID(I) And FlightNum >= 13 And FlightNum <= 16 Then
                Found = True
            End If
        Loop
        
         If (Not Found) Then
            MsgBox "Flight " & FlightNum & " is not a valid flight to " & Location & " Please Try Again", vbExclamation
            txtFlight.Text = ""
            
        Else
            MsgBox "You have selected flight number " & FlightID(I) & " which is the flight from " & TakeOff(I) & "."
            GoNext = True
            FlightTotal = Tickets * TicketCost(I)
            MsgBox "The total cost for all " & Tickets & " tickets that you have purchased for your flight from " & TakeOff(I) & " is " & FormatCurrency(FlightTotal)
        End If
    End If
    
    'saves data for three public variables to be used in the last checkout form
    Flight = FlightID(I)
    Leaving = TakeOff(I)
    Ticket = TicketCost(I)
    'running total of CheckoutTotal
    CheckoutTotal = CheckoutTotal + FlightTotal
    
End Sub

Private Sub cmdNextActivityPage_Click()
 'If/Next statement used for going to the activity page based on the location that the user chose earlier in the program
 'it will hide the current page and then go to whichever activity page that is associated with the location that was
 'saved earlier in the program
 
    If Location = "St. Joe" And GoNext = True Then
            frmFlight.Hide
            frmActivitiesJoseph.Show
        ElseIf Location = "Normandy" Then
            frmFlight.Hide
            frmActivitiesNormandy.Show
        ElseIf Location = "Badlands" Then
            frmFlight.Hide
            frmBadlandsActivity.Show
        ElseIf Location = "Saskatchewan" Then
            frmFlight.Hide
            frmCanada.Show
    Else
        MsgBox "Your flight does not end at your requested vacationing spot", vbExclamation
    End If
    
End Sub

Private Sub cmdQuit_Click()
    End     'ends the program when the user clicks this button
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading
 
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

    GoNext = False
End Sub
