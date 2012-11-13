VERSION 5.00
Begin VB.Form frmTotalCost 
   BackColor       =   &H000000C0&
   Caption         =   "Total Cost of Trip"
   ClientHeight    =   5550
   ClientLeft      =   1290
   ClientTop       =   1080
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   3345
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   6480
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "See the Cost of Trips for Other Cities"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdActualCost 
      Caption         =   "Compute the Actual Cost of the Trip"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picResultsActual 
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   3000
      Width           =   5775
   End
   Begin VB.PictureBox picResultsPref 
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   1680
      Width           =   5775
   End
   Begin VB.CommandButton cmdTotalCost 
      Caption         =   "Compute the Preferred Total Cost of the Trip"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Total cost of your TravelCity Trip"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   825
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmTotalCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'The purpose of this form is to finally calculate the total preferred cost of
                'the trip and the actual cost of the trip based on USAtourist. To calculate the
                'preferred cost of the trip the program just adds together the total cost from
                'frmTravel and frmOtherExp. To calculate the actual cost the program must first distingish
                'the characteristics of the trip and then use the values from various arrays and add them together
                
Option Explicit

Private Sub cmdActualCost_Click()
    'This button will calculate the actual cost of the trip based on USAtourist.com using a nested if,
    'and load an array with prices for a hotel and flight
    
    Open App.Path & "\Cities_Flight_Hotel.txt" For Input As #1
    
    CTR = 0
    
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Cities(CTR), TravelFlightOne(CTR), TravelFlightTwo(CTR), TravelFlightFour(CTR)
    Loop
    
    Close #1
    
    picResultsActual.Cls
    
    'These if statements find the characteristics of the trip based on user input
    If TravelerNumber = 1 And FlyCost > 0 And RentalCarCost > 0 Then 'This if statement will see how many travelers there are if there is fly cost and if there is rental cost. If there is not a match it will move on to another if statement.
        ActualTotalCost = TravelFlightOne(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = TravelFlightTwo(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = TravelFlightFour(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 1 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightOne(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightTwo(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightFour(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf DriveCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = DriveCost + (Hotels(CityNum) * 7) + RentalCar(CityNum) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf DriveCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = DriveCost + (Hotels(CityNum) * 7) + TotalOtherExp
        picResultsActual.Print Traveler; ", the actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    Else
        MsgBox "You have forgotten to compute part of your trip", , "Error!"
    End If

End Sub

Private Sub cmdBack_Click()
    'This button will bring the user back to frmUSATravel
    
    frmOtherExp.Visible = True
    frmTotalCost.Visible = False
    
End Sub

Private Sub cmdExit_Click()
    'This button will end the program
    End
End Sub

Private Sub cmdOther_Click()
    'This button will go to the frmOtherCities
    
    frmOtherCities.Visible = True
    frmTotalCost.Visible = False
    
End Sub

Private Sub cmdTotalCost_Click()
    'This button will add up all of the preferred costs for the trip to get a total
    'preferred cost
    
    picResultsPref.Cls
    
    PrefTotalCost = TotalTravelExp + TotalOtherExp 'Adds up the preferred price of the trip
    
    picResultsPref.Print Traveler; ", your total preferred cost for the trip to "; TravelCity; " is "; FormatCurrency(PrefTotalCost)
    
    
End Sub

Private Sub Form_Load()
    'When the form loads if will display the travelcity based on the user input in the label
    lblTotalCost.Caption = "Total Cost of your " & TravelCity & " Trip"
End Sub
