VERSION 5.00
Begin VB.Form frmTravel 
   BackColor       =   &H000000C0&
   Caption         =   "Travel Expenses"
   ClientHeight    =   7740
   ClientLeft      =   1080
   ClientTop       =   870
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   9900
   Begin VB.PictureBox picFly_Drive 
      Height          =   1215
      Left            =   7680
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.PictureBox picCarRental 
      Height          =   1215
      Left            =   7680
      Picture         =   "frmTravel.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   14
      Top             =   3720
      Width           =   2055
   End
   Begin VB.PictureBox picHotel 
      Height          =   1215
      Left            =   7680
      Picture         =   "frmTravel.frx":1337
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Compute Other Travel Expenses"
      Height          =   975
      Left            =   2760
      TabIndex        =   12
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Compute Total Expenses"
      Height          =   975
      Left            =   360
      TabIndex        =   11
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   5160
      TabIndex        =   10
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   7560
      TabIndex        =   9
      Top             =   6360
      Width           =   2055
   End
   Begin VB.PictureBox picResultsTotal 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3600
      ScaleHeight     =   795
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   5160
      Width           =   5895
   End
   Begin VB.PictureBox picResultsFly_Drive 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2520
      ScaleHeight     =   915
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   1200
      Width           =   4935
   End
   Begin VB.PictureBox picResultsCarRental 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2520
      ScaleHeight     =   915
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   3840
      Width           =   4935
   End
   Begin VB.PictureBox picResultsHotel 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2520
      ScaleHeight     =   915
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
   End
   Begin VB.CommandButton cmdCarRental 
      Caption         =   "Car Rental"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdHotel 
      Caption         =   "Hotel"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdFly_Drive 
      Caption         =   "Fly or Drive"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblTotalTravel 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Total Travel Expenses"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label lblTravel 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "TravelCity Travel Expenses"
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
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmTravel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'The purpose of this form is to calculate the main travel expenses. The user'
                'is able to decide whether to fly of drive and how many people to travel with.
                'Then the user can say what their preferred cost of a hotel is and whether they
                'want to rent a car. After all the costs have been calculated they can be added
                'up to get a total cost.

Option Explicit
Dim FlyDrive As String
Dim Rental As Single




Private Sub cmdBack_Click()
    'This button will bring the user back to frmCity
    
    frmTravel.Visible = False
    frmCity.Visible = True
    
End Sub

Private Sub cmdCarRental_Click()
    'This button will first ask the user if they wish to rent a car. If they do wish,
    'to rent a car it will calculate the preferred and actual cost of renting a car
    
    Open App.Path & "\Cities_RentalCar.txt" For Input As #3 'opens the file Cities_RentalCar.txt
    
    CTR = 0
    
    Do Until EOF(3) 'This do loop will load the file into the Cities and RentalCar Arrays
        CTR = CTR + 1
        Input #3, Cities(CTR), RentalCar(CTR)
    Loop
    
    Close #3 'Close the file
    
    picResultsCarRental.Cls 'Clear the contents of the picture
    
    Rental = InputBox("Do you wish to rent a car when you arrive to your travel location?, enter 1=Yes or 2=No", "Rental Car")
    
    Select Case Rental 'This select case will decide if the user has chosen to rent a car or not
        Case Is = 1 'if the user has chosen to rent a car then an inputbox asks for the preferred price
            PrefRentalCarCost = InputBox("What would your preferred price for a rental car be", "Preferred Cost")
            picResultsCarRental.Print "Your preferred cost for a rental car is "; FormatCurrency(PrefRentalCarCost)
            picResultsCarRental.Print "Your actual cost for a rental car in "; TravelCity; " is "; FormatCurrency(RentalCar(CityNum))
        Case Is = 2 'If the user doesn't want to rent a car then rental car cost is zero
            PrefRentalCarCost = 0
            picResultsCarRental.Print "Your estimated cost of a rental car is "; FormatCurrency(PrefRentalCarCost)
        Case Else
            MsgBox "You have entered an invalid value", , "Error"
    End Select
    
    
End Sub

Private Sub cmdExit_Click()
    End 'ends the program
End Sub

Private Sub cmdFly_Drive_Click()
    'This button will ask the user whether they want to fly to their vacation spot,
    'and if they do the button will ask the user what their preferred cost of flying
    'would be. This button will also load the array for cities, and miles to travel
    Dim GasPrice As Single
    Dim CarMileage As Single
    
    
    Open App.Path & "\Cities_Miles.txt" For Input As #1
    
    CTR = 0
    
    Do Until EOF(1) 'This do loop loads the file into the arrays cities and miles
        CTR = CTR + 1
        Input #1, Cities(CTR), Miles(CTR) 'stores the value of the arrays using CTR
    Loop

    Close #1 'closes the file
    
    
    
    picResultsFly_Drive.Cls
    
    FlyDrive = InputBox("Do you wish to fly or drive to your travel location, enter 1=fly or 2=drive", "Fly or Drive")
    
    Select Case FlyDrive 'This select case decides if the user wants to fly or drive and how many travelers there are
        Case Is = 1 'This case calculates the users cost for flying
            picFly_Drive = LoadPicture(App.Path & "\Pictures\plane.jpg")
            TravelerNumber = InputBox("How many people do you plan on coming on the trip, 1,2 or 4", "Traveler Number")
            Do Until TravelerNumber = 4 Or TravelerNumber = 2 Or TravelerNumber = 1 'This do loop sees if the user entered a valid number for travelers
                MsgBox "Your have entered an invalid number", , "error"
                TravelerNumber = InputBox("How many people do you plan on coming on the trip, 1,2 or 4", "Traveler Number")
            Loop
            FlyCost = InputBox("What would your preferred price per traveler for a flight be", "Preferred Cost")
            FlyCost = FlyCost * TravelerNumber 'multiplies preferred fly cost by travelernumber
            picResultsFly_Drive.Print "Your preferred cost for flying is "; FormatCurrency(FlyCost)
            DriveCost = 0 'makes drivecost zero
            picResultsFly_Drive.Print "Your preferred cost for driving is "; FormatCurrency(DriveCost)
        Case Is = 2 'This case calculates the cost for driving
            picFly_Drive = LoadPicture(App.Path & "\Pictures\car.jpg")
            TravelerNumber = InputBox("How many people do you plan on coming on the trip, 1,2 or 4", "Traveler Number")
            Do Until TravelerNumber = 4 Or TravelerNumber = 2 Or TravelerNumber = 1 'This do loop sees if the user entered a valid number for travelers
                MsgBox "Your have entered an invalid number", , "error"
                TravelerNumber = InputBox("How many people do you plan on coming on the trip, 1,2 or 4", "Traveler Number")
            Loop
            MsgBox "The distance to " & TravelCity & " from Minneapolis is " & Miles(CityNum) & " miles"
            GasPrice = InputBox("How much to you estimate the cost of gas per gallon will be", "Gas Cost")
            CarMileage = InputBox("What is the estimated miles to the gallon for your car", "Car Mileage")
            DriveCost = (Miles(CityNum) / CarMileage) * GasPrice 'calculates cost of driving
            picResultsFly_Drive.Print "Your estimated cost of driving is "; FormatCurrency(DriveCost)
            FlyCost = 0
            picResultsFly_Drive.Print "Your preferred cost for flying is "; FormatCurrency(FlyCost)
        Case Else
            MsgBox "Your have entered an invalid value", , "Error"
    End Select
    
End Sub

Private Sub cmdHotel_Click()
    'This button will calculate the preferred cost of a hotel for a week in the travel,
    'City. It will also load the array for cost of hotels
    
    Open App.Path & "\Cities_Hotel.txt" For Input As #2
    
    CTR = 0
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, Cities(CTR), Hotels(CTR)
    Loop
    
    Close #2
    
    picResultsHotel.Cls
    
    PrefHotelCost = InputBox("What would you prefer to pay per night for a hotel room for a week long vacation in " & TravelCity)
    PrefHotelCost = PrefHotelCost * 7 'Calculates the cost of a hotel based on a week long stay
    HotelCost = Hotels(CityNum) * 7
    
    picResultsHotel.Print "Your preferred cost for a hotel room for a week is "; FormatCurrency(PrefHotelCost)
    picResultsHotel.Print "The actual cost for a hotel room in "; TravelCity
    picResultsHotel.Print "for a week is "; FormatCurrency(HotelCost)
        
End Sub

Private Sub cmdOther_Click()
    'This button will bring the user to the frmOtherExp if the user has caclulated the '
    'total cost
    
    If TotalTravelExp > 0 Then
        frmOtherExp.Visible = True
        frmTravel.Visible = False
    Else
        MsgBox "Please calculate total travel expenses.", , "Error"
    End If
    
    
End Sub

Private Sub cmdTotal_Click()
    'this button will compute the total preferred travel expenses
    
    picResultsTotal.Cls
    
    TotalTravelExp = PrefHotelCost + FlyCost + DriveCost + PrefRentalCarCost 'Calculates the total travel costs based on previous entries
    
    picResultsTotal.Print "The total travel expenses are "; FormatCurrency(TotalTravelExp)
    
End Sub


Private Sub Form_Load()
    'When the form loads it will display in the caption of lblTravel the selected
    'travel city along with travel expenses
    
    lblTravel.Caption = TravelCity & " Travel Expenses"
    
End Sub


