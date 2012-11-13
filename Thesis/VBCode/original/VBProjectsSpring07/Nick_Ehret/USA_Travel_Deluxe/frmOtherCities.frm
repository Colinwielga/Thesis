VERSION 5.00
Begin VB.Form frmOtherCities 
   BackColor       =   &H000000C0&
   Caption         =   "Other Cities"
   ClientHeight    =   7530
   ClientLeft      =   1080
   ClientTop       =   1080
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7980
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Click to See the Final Page!"
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   2963
      TabIndex        =   7
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   5520
      TabIndex        =   6
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdLeastMost 
      Caption         =   "Click to Sort the List from Least Expensive to Most Expensive"
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdMaxMin 
      Caption         =   "Click to Sort the List from Most Expensive to Least Expensive"
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdMatch 
      Caption         =   "Click to Show a Specific City's Trip Cost"
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdOtherCities 
      Caption         =   "Costs for Trips to Other Cities"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.PictureBox picResultsCities 
      Height          =   4935
      Left            =   2880
      ScaleHeight     =   4875
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblOtherCities 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Other City Trip Expenses"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmOtherCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'The purpose of this form is to display the actual costs of traveling to other'
                'cities based on the characteristics of a trip from user inputs. This form will also'
                'sort the list of costs from most expensive to least expensive and vice versa. This form will
                'also allow the user to search for a specific cost based on the city name
                
Option Explicit

Private Sub cmdBack_Click()
    'This button will bring the user back to frmTotalCost
    
    frmTotalCost.Visible = True
    frmOtherCities.Visible = False
    
End Sub

Private Sub cmdCredits_Click()
    'this button will bring the user to frmFinish
    
    frmFinish.Visible = True
    frmOtherCities.Visible = False
    
End Sub

Private Sub cmdExit_Click()
    'This button will end the program
    End
End Sub

Private Sub cmdLeastMost_Click()
    'This button will sort the list of cities from least expensive to most expensive using
    'a bubble sort method
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempCost As Single
    Dim TempCity As String
    Dim Pos As Integer
    
    picResultsCities.Cls
    
    For Pass = 1 To (CTR - 1)
        For Comp = 1 To (CTR - Pass)
            If ActualCityCost(Comp) > ActualCityCost(Comp + 1) Then 'If the city cost is bigger this will switch the values
                TempCost = ActualCityCost(Comp)
                ActualCityCost(Comp) = ActualCityCost(Comp + 1)
                ActualCityCost(Comp + 1) = TempCost
                
                TempCity = Cities(Comp) 'This brings the name of the city along with the cost
                Cities(Comp) = Cities(Comp + 1)
                Cities(Comp + 1) = TempCity
            End If
        Next Comp
    Next Pass
    
    picResultsCities.Print "Cities", Tab(30); "Cost of the Trip" 'this will make the display look better
    picResultsCities.Print "********************************************************"
    
    Pos = 0
    For Pos = 1 To CTR 'This will display all the costs of the city and the cities name after they have been sorted
        picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
    Next Pos

End Sub

Private Sub cmdMatch_Click()
    'This button will match the user with a certain city which is inputed by the user
    Dim Found As Boolean
    Dim Pos As Integer
    Dim Search As String
    
    
    picResultsCities.Cls
    
    Found = False
    Pos = 0
     
    'Ask the user for a city name
    Search = InputBox("Please Enter a City in which you would like to see the actual cost", "Search")
    
    Do While (Found = False And Pos < CTR) 'finds the serch value based on the input from the user
        Pos = Pos + 1
        If LCase(Search) = LCase(Cities(Pos)) Then 'allows the search value to be lower case letters
            Found = True
        End If
    Loop
    
    picResultsCities.Cls
    
    picResultsCities.Print "Cities", Tab(30); "Cost of the Trip"
    picResultsCities.Print "********************************************************"
    
    If Found = True Then 'if a match is found then it displays the match and if not an error message is displayed
        picResultsCities.Print Cities(Pos); Tab(30); FormatCurrency(ActualCityCost(Pos))
    Else
        MsgBox "You have not entered a city from the list", , "Error"
    End If
    
        
        
    
End Sub

Private Sub cmdMaxMin_Click()
    'This button will sort the list from most expensive to least expensive using a bubble
    'sort method
    Dim Pass As Integer
    Dim Comp As Integer
    Dim TempCost As Single
    Dim TempCity As String
    Dim Pos As Integer
    
    picResultsCities.Cls
    
    For Pass = 1 To (CTR - 1)
        For Comp = 1 To (CTR - Pass)
            If ActualCityCost(Comp) < ActualCityCost(Comp + 1) Then
                TempCost = ActualCityCost(Comp)
                ActualCityCost(Comp) = ActualCityCost(Comp + 1)
                ActualCityCost(Comp + 1) = TempCost
                
                TempCity = Cities(Comp) 'brings the city names along with the costs
                Cities(Comp) = Cities(Comp + 1)
                Cities(Comp + 1) = TempCity
            End If
        Next Comp
    Next Pass
    
    picResultsCities.Print "Cities", Tab(30); "Cost of the Trip"
    picResultsCities.Print "********************************************************"
    
    Pos = 0
    For Pos = 1 To CTR 'Prints the sorted list
        picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
    Next Pos
    
            
End Sub

Private Sub cmdOtherCities_Click()
    'This button will display the cost of traveling to other cities in a formatted picture
    'box
    Dim Pos As Integer
    
    picResultsCities.Cls
    
    picResultsCities.Print "Cities", Tab(30); "Cost of the Trip"
    picResultsCities.Print "********************************************************"
    
    
    Pos = 0
    
    'This if statement will find the actual cost based on user specifications and also load that actual cost into an array
    If TravelerNumber = 1 And FlyCost > 0 And RentalCarCost > 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightOne(Pos) + RentalCar(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost > 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightTwo(Pos) + RentalCar(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost > 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightFour(Pos) + RentalCar(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf TravelerNumber = 1 And FlyCost > 0 And RentalCarCost = 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightOne(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost = 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightTwo(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost = 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = TravelFlightFour(Pos) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf DriveCost > 0 And RentalCarCost > 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = DriveCost + (Hotels(Pos) * 7) + RentalCar(CityNum) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    ElseIf DriveCost > 0 And RentalCarCost = 0 Then
        For Pos = 1 To CTR
            ActualCityCost(Pos) = DriveCost + (Hotels(Pos) * 7) + TotalOtherExp
            picResultsCities.Print Cities(Pos), Tab(30); FormatCurrency(ActualCityCost(Pos))
        Next Pos
    Else
        MsgBox "Your have forgotten to compute part of your trip", , "Error"
    End If
    
End Sub

