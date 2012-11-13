VERSION 5.00
Begin VB.Form frmFlorida 
   BackColor       =   &H0080FFFF&
   Caption         =   "Visit Florida"
   ClientHeight    =   8775
   ClientLeft      =   2220
   ClientTop       =   1470
   ClientWidth     =   11055
   LinkTopic       =   "Form5"
   ScaleHeight     =   8775
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   3375
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFF80&
      Height          =   3375
      Left            =   4320
      ScaleHeight     =   3315
      ScaleWidth      =   6195
      TabIndex        =   5
      Top             =   5160
      Width           =   6255
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H00FFFF80&
      Height          =   2895
      Left            =   5160
      Picture         =   "frmFlorida.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton cmdCarRental 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Car Rentals"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlane 
      BackColor       =   &H0080FF80&
      Caption         =   "Plane Tickets"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdActivites 
      BackColor       =   &H0080C0FF&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdHotel 
      BackColor       =   &H00FF80FF&
      Caption         =   "Hotels"
      BeginProperty Font 
         Name            =   "Minion Pro Med"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblFloridaTitle 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Orlando, Florida"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   8775
   End
End
Attribute VB_Name = "frmFlorida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: Florida
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform the viewers of the information that is in Orlando
'           when making a decision on travel destination.

'Makes the user declare all the variables
Option Explicit

Private Sub cmdActivites_Click()
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
'Sets ctr to 0
sum = 0

'Loads the files so that it can be read by the user.
Open App.Path & "\OrlandoActivities.txt" For Input As #3
Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, activities(ctr), price(ctr)
Loop

'Loads and shows a picture of an activity that is available in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\OrlandoActivities.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Activities that are avialable in the Orlando area include:"
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
Oave = sum / ctr

'Prints the average activity price.
picResults2.Print
picResults2.Print "The average price of an activity is:"; Tab(45); FormatCurrency(Oave, 0)

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
Open App.Path & "\OrlandoCarRental.txt" For Input As #1
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
OAvisavg = sumAvis / ctr
OAlamoavg = sumAlamo / ctr
ONationalavg = sumNational / ctr

'Prints the averages.
picResults2.Print
picResults2.Print "Averages:", FormatCurrency(OAvisavg, 0), FormatCurrency(OAlamoavg, 0), FormatCurrency(ONationalavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #1

End Sub

Private Sub cmdHome_Click()
'When button is clicked, the Home Page shows and the Orlando form hides.
frmHome.Show
frmFlorida.Hide
End Sub

Private Sub cmdHotel_Click()
'Declare all the variables needed for this command.
Dim hotels(1 To 11) As String
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
Open App.Path & "\OrlandoHotels.txt" For Input As #2
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, hotels(ctr), price(ctr)
Loop

'Loads and shows a picture of a Hotel that is offered in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\OrlandoHotel.jpg")

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
Oavg = sum / ctr

'Prints the average hotel price.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(45); FormatCurrency(Oavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #2

End Sub

Private Sub cmdPlane_Click()
'Declare all the variables needed for this command.
Dim Oflights(1 To 6) As String
Dim Oprice(1 To 2000) As Integer
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
Open App.Path & "\OrlandoFlights.txt" For Input As #4
Do While Not EOF(4)
    ctr = ctr + 1
    Input #4, Oflights(ctr), Oprice(ctr)
Loop

'Loads and shows a picture of a plane that you could fly on to get to your destination.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Airplane4.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Airline Flights that are avialable to Orlando include:"
picResults2.Print
picResults2.Print "Airline"; Tab(40); "Price"
picResults2.Print "-----------------------------------------------------------------------------"

'Puts the information in alphabetical order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Oflights(pos) > Oflights(pos + 1) Then
            tempflights = Oflights(pos)
            Oflights(pos) = Oflights(pos + 1)
            Oflights(pos + 1) = tempflights
            tempprice = Oprice(pos)
            Oprice(pos) = Oprice(pos + 1)
            Oprice(pos + 1) = tempprice
        End If
    Next pos
Next pass

'Prints the information.
For y = 1 To ctr
    picResults2.Print Oflights(y); Tab(40); FormatCurrency(Oprice(y), 0)
    sum = sum + Oprice(y)
Next y

'Finds the average flight cost.
OPave = sum / ctr

'Prints the average flight cost.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(40); FormatCurrency(OPave, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #4

End Sub
