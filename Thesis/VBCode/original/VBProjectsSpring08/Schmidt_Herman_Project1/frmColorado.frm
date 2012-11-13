VERSION 5.00
Begin VB.Form frmColorado 
   BackColor       =   &H00008000&
   Caption         =   "Visit Colorado"
   ClientHeight    =   9630
   ClientLeft      =   1305
   ClientTop       =   1125
   ClientWidth     =   12570
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form6"
   ScaleHeight     =   9630
   ScaleWidth      =   12570
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H000080FF&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8640
      Width           =   3135
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   6000
      ScaleHeight     =   3435
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   3000
      Width           =   5655
   End
   Begin VB.PictureBox picResults1 
      Height          =   2895
      Left            =   600
      Picture         =   "frmColorado.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   4
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdCarRental 
      BackColor       =   &H000080FF&
      Caption         =   "Car Rentals"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdActivities 
      BackColor       =   &H000080FF&
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlane 
      BackColor       =   &H000080FF&
      Caption         =   "Plane Tickets"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdHotel 
      BackColor       =   &H000080FF&
      Caption         =   "Hotels"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label lblDenver 
      BackColor       =   &H00800000&
      Caption         =   "Denver, Colorado"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   8175
   End
End
Attribute VB_Name = "frmColorado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: Colorado
'Author: Taylor Herman & Mindy Schmidt
'Date Written: 3/23/08
'Objective: To inform the viewers of the information that is in Denver
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
Open App.Path & "\DenverActivities.txt" For Input As #3
Do While Not EOF(3)
    ctr = ctr + 1
    Input #3, activities(ctr), price(ctr)
Loop

'Loads and shows a picture of an activity that is available in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\DenverActivities.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Activities that are avialable in the Denver area include:"
picResults2.Print
picResults2.Print "Activities"; Tab(50); "Price"
picResults2.Print "----------------------------------------------------------------------------------------------------"

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
    picResults2.Print activities(H); Tab(50); FormatCurrency(price(H), 0)
    sum = sum + price(H)
Next H

'Finds the average activity price.
Dave = sum / ctr

'Prints the average activity price.
picResults2.Print
picResults2.Print "The average price of an activity is:"; Tab(50); FormatCurrency(Dave, 0)

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
Open App.Path & "\DenverCarRental.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, cars(ctr), priceAvis(ctr), priceAlamo(ctr), priceNational(ctr)
Loop

'Loads and shows a picture of a car that is offered by the Car Rental Companies.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Car3.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Here are the choices of Car Rental Companies in the area:"
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
DAvisavg = sumAvis / ctr
DAlamoavg = sumAlamo / ctr
DNationalavg = sumNational / ctr

'Prints the averages.
picResults2.Print
picResults2.Print "Averages:", FormatCurrency(DAvisavg, 0), FormatCurrency(DAlamoavg, 0), FormatCurrency(DNationalavg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #1

End Sub

Private Sub cmdHome_Click()
'When button is clicked, the Home Page shows and the Colorado form hides.
frmHome.Show
frmColorado.Hide
End Sub

Private Sub cmdHotel_Click()
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
Open App.Path & "\DenverHotels.txt" For Input As #2
Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, hotels(ctr), price(ctr)
Loop

'Loads and shows a picture of a Hotel that is offered in the area.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\DenverHotel.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Here are the choices of Hotels in the area:"
picResults2.Print
picResults2.Print "Hotels"; Tab(55); "Price"
picResults2.Print "-------------------------------------------------------------------------------------------------------------"

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
    picResults2.Print hotels(i); Tab(55); FormatCurrency(price(i), 0)
    sum = sum + price(i)
Next i

'Finds the average hotel price.
Davg = sum / ctr

'Prints the average hotel price.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(55); FormatCurrency(Davg, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #2

End Sub

Private Sub cmdPlane_Click()
'Declare all the variables needed for this command.
Dim Dflights(1 To 10) As String
Dim Dprice(1 To 2000) As Integer
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
Open App.Path & "\DenverFlights.txt" For Input As #4
Do While Not EOF(4)
    ctr = ctr + 1
    Input #4, Dflights(ctr), Dprice(ctr)
Loop

'Loads and shows a picture of a plane that you could fly on to get to your destination.
picResults1.Picture = LoadPicture(App.Path & "\ProjectPictures\Airplane3.jpg")

'Informs the user of what is going to be displayed below.
picResults2.Cls
picResults2.Print "Airline Flights that are avialable to Denver include:"
picResults2.Print
picResults2.Print "Airline"; Tab(40); "Price"
picResults2.Print "----------------------------------------------------------------------------"

'Puts the information in alphabetical order.
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Dflights(pos) > Dflights(pos + 1) Then
            tempflights = Dflights(pos)
            Dflights(pos) = Dflights(pos + 1)
            Dflights(pos + 1) = tempflights
            tempprice = Dprice(pos)
            Dprice(pos) = Dprice(pos + 1)
            Dprice(pos + 1) = tempprice
        End If
    Next pos
Next pass

'Prints the information.
For y = 1 To ctr
    picResults2.Print Dflights(y); Tab(40); FormatCurrency(Dprice(y), 0)
    sum = sum + Dprice(y)
Next y

'Finds the average flight cost.
DPave = sum / ctr

'Prints the average flight cost.
picResults2.Print
picResults2.Print "The average price of a hotel is:"; Tab(40); FormatCurrency(DPave, 0)

'Closes the the file so that it can be opened again when clicked on.
Close #4

End Sub
