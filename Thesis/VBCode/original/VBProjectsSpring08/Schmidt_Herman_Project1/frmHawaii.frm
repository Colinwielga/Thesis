VERSION 5.00
Begin VB.Form frmHawaii 
   BackColor       =   &H008080FF&
   Caption         =   "Visit Hawaii"
   ClientHeight    =   9900
   ClientLeft      =   1875
   ClientTop       =   1020
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   Picture         =   "frmHawaii.frx":0000
   ScaleHeight     =   9900
   ScaleWidth      =   11685
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   2775
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   5280
      ScaleHeight     =   3675
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   5160
      Width           =   5655
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5760
      Picture         =   "frmHawaii.frx":5400
      ScaleHeight     =   2955
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton cmdActivities 
      BackColor       =   &H0080C0FF&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlane 
      BackColor       =   &H0080C0FF&
      Caption         =   "Plane Tickets"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdCarRental 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Car Rental"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdHotel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Hotels"
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblHawaiiTitle 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Honolulu, Hawaii"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmHawaii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: California
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform the viewers of the information that is in Honolulu
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
Dim sum As Integer, ave As Integer
Dim ctr As Integer
Dim pass As Integer, pos As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum to 0
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\HonoluluActivities.txt" For Input As #3
Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, activities(ctr), price(ctr)
Loop

'Loads and shows a picture of an activity that is available in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\HonoluluActivities.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Activities that are avialable in the Honolulu area include:"
picResults2.Print
picResults2.Print "Activities"; Tab(66); "Price"
picResults2.Print "---------------------------------------------------------------------------------------------------------------------"

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
    picResults2.Print activities(H); Tab(66); FormatCurrency(price(H), 0)
    sum = sum + price(H)
Next H

'Finds the average activity price.
Have = sum / ctr

'Prints the average activity price.
picResults2.Print
picResults2.Print "The average price of an activity is:"; Tab(66); FormatCurrency(Have, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #3

End Sub

Private Sub cmdCarRental_Click()
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
Open App.Path & "\HonoluluCarRental.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, cars(ctr), priceAvis(ctr), priceAlamo(ctr), priceNational(ctr)
Loop

'Loads and shows a picture of a car that is offered by the Car Rental Companies.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Car4.jpg")

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
HAvisavg = sumAvis / ctr
HAlamoavg = sumAlamo / ctr
HNationalavg = sumNational / ctr

'Prints the averages.
picResults2.Print
picResults2.Print "Averages:", FormatCurrency(HAvisavg, 0), FormatCurrency(HAlamoavg, 0), FormatCurrency(HNationalavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #1

End Sub

Private Sub cmdHome_Click()
'When button is clicked, the Home Page shows and the Honolulu form hides.
frmHome.Show
frmHawaii.Hide
End Sub


Private Sub cmdHotel_Click()
'Declare all the variables needed for this command.
Dim hotels(1 To 10) As String
Dim price(1 To 300) As Integer
Dim i As Integer
Dim temphotels As String
Dim tempprice As Integer
Dim sum As Integer
Dim avg As Single
Dim ctr As Integer
Dim pass As Integer, pos As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum to 0.
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\HonoluluHotels.txt" For Input As #2
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, hotels(ctr), price(ctr)
Loop

'Loads and shows a picture of a Hotel that is offered in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\HonoluluHotel.jpg")

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
Havg = sum / ctr

'Prints the average hotel price.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(45); FormatCurrency(Havg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #2

End Sub

Private Sub cmdPlane_Click()
'Declare all the variables needed for this command.
Dim Hflights(1 To 6) As String
Dim Hprice(1 To 2000) As Integer
Dim y As Integer
Dim tempflights As String
Dim tempprice As Integer
Dim ctr As Integer
Dim pass As Integer, pos As Integer
Dim sum As Integer

'Sets ctr to 0.
ctr = 0
'Sets sum to 0
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\HonoluluFlights.txt" For Input As #4
Do While Not EOF(4)
    ctr = ctr + 1
    Input #4, Hflights(ctr), Hprice(ctr)
Loop

'Loads and shows a picture of a plane that you could fly on to get to your destination.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Airplane5.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Airline Flights that are avialable to Honolulu include:"
picResults2.Print
picResults2.Print "Airline"; Tab(40); "Price"
picResults2.Print "-----------------------------------------------------------------------------"

'Puts the information in alphabetical order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Hflights(pos) > Hflights(pos + 1) Then
            tempflights = Hflights(pos)
            Hflights(pos) = Hflights(pos + 1)
            Hflights(pos + 1) = tempflights
            tempprice = Hprice(pos)
            Hprice(pos) = Hprice(pos + 1)
            Hprice(pos + 1) = tempprice
        End If
    Next pos
Next pass

'Prints the information.
For y = 1 To ctr
    picResults2.Print Hflights(y); Tab(40); FormatCurrency(Hprice(y), 0)
    sum = sum + Hprice(y)
Next y

'Finds the average flight cost.
HPave = sum / ctr

'Prints the average flight cost.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(40); FormatCurrency(HPave, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #4

End Sub

