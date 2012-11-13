VERSION 5.00
Begin VB.Form frmCalifornia 
   BackColor       =   &H000080FF&
   Caption         =   "Visit California"
   ClientHeight    =   8820
   ClientLeft      =   1995
   ClientTop       =   1575
   ClientWidth     =   11190
   FillColor       =   &H0000FFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   ScaleHeight     =   8820
   ScaleWidth      =   11190
   Begin VB.PictureBox picResults1 
      FillColor       =   &H0080FFFF&
      Height          =   2655
      Left            =   4800
      Picture         =   "frmCaliforniafrm.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H0080FFFF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   2775
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H0080FFFF&
      Height          =   3495
      Left            =   3720
      ScaleHeight     =   3435
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   4800
      Width           =   6495
   End
   Begin VB.CommandButton cmdPlane 
      BackColor       =   &H0000FFFF&
      Caption         =   "Plane Tickets"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmdActivities 
      BackColor       =   &H0000FFFF&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdHotels 
      BackColor       =   &H0000FFFF&
      Caption         =   "Hotels"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdCarRentals 
      BackColor       =   &H0000FFFF&
      Caption         =   "Car Rentals"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblCaliforniaTitle 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "San Diego, California"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmCalifornia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: California
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform the viewers of the information that is in San Diego
'           when making a decision on travel destination.

'Makes the user declare all the variables
Option Explicit

Private Sub cmdActivities_Click()
'Declare all the variables needed for this command.
Dim activities(1 To 10) As String
Dim price(1 To 200) As Integer
Dim H As Integer
Dim tempactivities As String
Dim tempprice As Integer
Dim sum As Integer
Dim ctr As Integer
Dim pass As Integer, pos As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum to 0
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\SanDiegoActivities.txt" For Input As #3
Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, activities(ctr), price(ctr)
Loop

'Loads and shows a picture of an activity that is available in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\SanDiegoActivities.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Activities that are avialable in the San Diego area include:"
picResults2.Print
picResults2.Print "Activities"; Tab(45); "Price"
picResults2.Print "------------------------------------------------------------------------------------"

'Puts the information in ascending order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If price(pos) > price(pos + 1) Then
            tempprice = price(pos)
            price(pos) = price(pos + 1)
            price(pos + 1) = tempprice
            tempactivities = activities(pos)
            activities(pos) = activities(pos + 1)
            activities(pos + 1) = tempactivities
        End If
    Next pos
Next pass

'Prints the information.
For H = 1 To ctr
    picResults2.Print activities(H); Tab(45); FormatCurrency(price(H), 0)
    sum = sum + price(H)
Next H

'Finds the average activity price.
SDave = sum / ctr

'Prints the average activity price.
picResults2.Print
picResults2.Print "The average price of an activity is:"; Tab(45); FormatCurrency(SDave, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #3

End Sub

Private Sub cmdCarRentals_Click()
'Declare all the variables needed for this command.
Dim cars(1 To 3) As String
Dim priceAvis(1 To 100) As Integer, priceAlamo(1 To 100) As Integer, priceNational(1 To 100) As Integer
Dim j As Integer
Dim ctr As Integer
Dim sumAvis As Integer
Dim sumAlamo As Integer
Dim sumNational As Integer

'Sets ctr to 0.
ctr = 0
'Sets sums to 0
sumAvis = 0
sumAlamo = 0
sumNational = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\SanDiegoCarRental.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, cars(ctr), priceAvis(ctr), priceAlamo(ctr), priceNational(ctr)
Loop

'Loads and shows a picture of a car that is offered by the Car Rental Companies.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Car2.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Here are the choices of Car Rental Comapanies in the area:"
picResults2.Print
picResults2.Print "Cars", "Avis", "Alamo", "National"
picResults2.Print "---------------------------------------------------------------------------------------"

'Prints the information.
For j = 1 To ctr
    picResults2.Print cars(j), FormatCurrency(priceAvis(j), 0), FormatCurrency(priceAlamo(j), 0), FormatCurrency(priceNational(j), 0)
    sumAvis = sumAvis + priceAvis(j)
    sumAlamo = sumAlamo + priceAlamo(j)
    sumNational = sumNational + priceNational(j)
Next j

'Finds the averages of the car rentals.
SDAvisavg = sumAvis / ctr
SDAlamoavg = sumAlamo / ctr
SDNationalavg = sumNational / ctr

'Prints the averages.
picResults2.Print
picResults2.Print "Averages:", FormatCurrency(SDAvisavg, 0), FormatCurrency(SDAlamoavg, 0), FormatCurrency(SDNationalavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #1

End Sub

Private Sub cmdHome_Click()
'When button is clicked, the Home Page shows and the California form hides.
frmHome.Show
frmCalifornia.Hide
End Sub

Private Sub cmdHotels_Click()
'Declare all the variables needed for this command.
Dim hotels(1 To 10) As String
Dim price(1 To 300) As Integer
Dim i As Integer
Dim temphotels As String
Dim tempprice As Integer
Dim sum As Integer
Dim ctr As Integer
Dim pass As Integer, pos As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum to 0.
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\SanDiegoHotels.txt" For Input As #2
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, hotels(ctr), price(ctr)
Loop

'Loads and shows a picture of a Hotel that is offered in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\SanDiegoHotel.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Here are the choices of Hotels in the area:"
picResults2.Print
picResults2.Print "Hotels"; Tab(45); "Price"
picResults2.Print "-------------------------------------------------------------------------------------"

'Puts the information in ascending order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If price(pos) > price(pos + 1) Then
            tempprice = price(pos)
            price(pos) = price(pos + 1)
            price(pos + 1) = tempprice
            temphotels = hotels(pos)
            hotels(pos) = hotels(pos + 1)
            hotels(pos + 1) = temphotels
        End If
    Next pos
Next pass

'Prints the information and finds the sum of the hotel prices for the average.
For i = 1 To ctr
    picResults2.Print hotels(i); Tab(45); FormatCurrency(price(i), 0)
    sum = sum + price(i)
Next i

'Finds the average hotel price.
SDavg = sum / ctr

'Prints the average hotel price.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(45); FormatCurrency(SDavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #2

End Sub

Private Sub cmdPlane_Click()
'Declare all the variables needed for this command.
Dim SDflights(1 To 10) As String
Dim SDprice(1 To 2000) As Integer
Dim y As Integer
Dim tempflights As String
Dim tempprice As Integer
Dim ctr As Integer
Dim pass As Integer, pos As Integer
Dim sum As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum = 0
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\SanDiegoFlights.txt" For Input As #4
Do While Not EOF(4)
    ctr = ctr + 1
    Input #4, SDflights(ctr), SDprice(ctr)
Loop

'Loads and shows a picture of a plane that you could fly on to get to your destination.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\AirPlane2.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Airline Flights that are avialable to San Diego include:"
picResults2.Print
picResults2.Print "Airline"; Tab(40); "Price"
picResults2.Print "------------------------------------------------------------------------------------"

'Puts the information in alphabetical order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If SDflights(pos) > SDflights(pos + 1) Then
            tempflights = SDflights(pos)
            SDflights(pos) = SDflights(pos + 1)
            SDflights(pos + 1) = tempflights
            tempprice = SDprice(pos)
            SDprice(pos) = SDprice(pos + 1)
            SDprice(pos + 1) = tempprice
        End If
    Next pos
Next pass

'Prints the information.
For y = 1 To ctr
    picResults2.Print SDflights(y); Tab(40); FormatCurrency(SDprice(y), 0)
    sum = sum + SDprice(y)
Next y

'Finds the average flight cost.
SDPave = sum / ctr

'Prints the average flight cost.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(40); FormatCurrency(SDPave, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #4

End Sub
