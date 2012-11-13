VERSION 5.00
Begin VB.Form frmFinish 
   BackColor       =   &H000000C0&
   Caption         =   "USA Travel"
   ClientHeight    =   6930
   ClientLeft      =   885
   ClientTop       =   1080
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9120
   Begin VB.PictureBox picResults 
      Height          =   3615
      Left            =   2520
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   1680
      Width           =   6255
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "Click to See a Picture of your Travel City"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdCost 
      Caption         =   "Click to See your Final Trip Costs Again"
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   975
      Left            =   6480
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "Click to See the Works Cited List"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblFinish 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Thank You For Using USA Travel Deluxe!"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1493
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: USA Travel Planner Deluxe
'Written By: Nick Ehret
'Date Written: March 21, 2007
'Form Purpose: 'This is the final form of the project and its purpose is to display a works
                'cited list, show the cost of the trips again, and to display a picture of the
                'travel city.
Option Explicit


Private Sub cmdBack_Click()
    'This button bring the user back to frmOtherCities
    
    frmOtherCities.Visible = True
    frmFinish.Visible = False
    
End Sub

Private Sub cmdCost_Click()
    'This button will display both the preferred and final trip costs again
    
    picResults.Cls
    
    picResults.Print "Costs of the trip to " & TravelCity
    picResults.Print "*********************************************************"
    
    picResults.Print Traveler; ", your total preferred cost for the trip to "; TravelCity; " is "; FormatCurrency(PrefTotalCost)
    
    picResults.Print
    
    'This will display the actual cost again just like in the previous forms
    If TravelerNumber = 1 And FlyCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = TravelFlightOne(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = TravelFlightTwo(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = TravelFlightFour(CityNum) + RentalCar(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 1 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightOne(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 2 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightTwo(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf TravelerNumber = 4 And FlyCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = TravelFlightFour(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf DriveCost > 0 And RentalCarCost > 0 Then
        ActualTotalCost = DriveCost + (Hotels(CityNum) * 7) + RentalCar(CityNum) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    ElseIf DriveCost > 0 And RentalCarCost = 0 Then
        ActualTotalCost = DriveCost + (Hotels(CityNum) * 7) + TotalOtherExp
        picResults.Print Traveler; ", your actual cost for your trip to "; TravelCity; " is "; FormatCurrency(ActualTotalCost)
    Else
        MsgBox "You have forgotten to compute part of your trip", , "Error!"
    End If

    
    
End Sub

Private Sub cmdExit_Click()
    'this program will end the program
    End
End Sub

Private Sub cmdPic_Click()
    'This button will load the picture of the travel city using select case and based,
    'on the citynum
    
    picResults.Cls
    
    
    Select Case CityNum
        Case 1
            picResults = LoadPicture(App.Path & "\Pictures\SanFrancisco.jpg")
        Case 2
            picResults = LoadPicture(App.Path & "\Pictures\NewYorkCity.jpg")
        Case 3
            picResults = LoadPicture(App.Path & "\Pictures\LosAngeles.jpg")
        Case 4
            picResults = LoadPicture(App.Path & "\Pictures\Chicago.jpg")
        Case 5
            picResults = LoadPicture(App.Path & "\Pictures\Orlando.jpg")
        Case 6
            picResults = LoadPicture(App.Path & "\Pictures\Miami.jpg")
        Case 7
            picResults = LoadPicture(App.Path & "\Pictures\Seattle.jpg")
        Case 8
            picResults = LoadPicture(App.Path & "\Pictures\NewOrleans.jpg")
        Case 9
            picResults = LoadPicture(App.Path & "\Pictures\Denver.jpg")
        Case 10
            picResults = LoadPicture(App.Path & "\Pictures\Nashville.jpg")
        Case 11
            picResults = LoadPicture(App.Path & "\Pictures\Boston.jpg")
        Case 12
            picResults = LoadPicture(App.Path & "\Pictures\WashingtonDC.jpg")
        Case 13
            picResults = LoadPicture(App.Path & "\Pictures\Atlanta.jpg")
        Case 14
            picResults = LoadPicture(App.Path & "\Pictures\Phoenix.jpg")
        Case 15
            picResults = LoadPicture(App.Path & "\Pictures\LasVegas.jpg")
        Case Else
            MsgBox "You have not chosen a city", , "Error"
    End Select
    
End Sub



Private Sub cmdWorksCited_Click()
    'This button will print a list of sources I used for this project
    picResults.Cls
    
    
    picResults.Print "Works Cited"; Tab(30); "Websites"
    picResults.Print "*************************************************************"
    
    picResults.Print "Travel Expenses"; Tab(30); "http://www.usatourist.com/"
    picResults.Print "Pictures"; Tab(30); "http://images.google.com/imghp?hl=en"
    
End Sub
