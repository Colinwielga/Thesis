VERSION 5.00
Begin VB.Form frmFlights 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   FillColor       =   &H00C0E0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd1 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7920
      TabIndex        =   10
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdBahamas 
      BackColor       =   &H0000FF00&
      Caption         =   "Bahamas"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdAtlanta 
      BackColor       =   &H000080FF&
      Caption         =   "Atlanta"
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdPanamaCity 
      BackColor       =   &H008080FF&
      Caption         =   "Panama City"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdMiami 
      BackColor       =   &H00C0C000&
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdLasVegas 
      BackColor       =   &H008080FF&
      Caption         =   "Las Vegas"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSanDiego 
      BackColor       =   &H00FFFFFF&
      Caption         =   "San Diego"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancun 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cancun"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton cmdPrices 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here FIRST For Flight Information"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   2460
      Left            =   480
      Picture         =   "frmFlights.frx":0000
      Top             =   1440
      Width           =   3645
   End
   Begin VB.Label lblPrices 
      Alignment       =   2  'Center
      Caption         =   "Flight Prices"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmFlights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flight Prices'
'This is the flight pricing page'
'Its all about the money'
'I order the flights from most to least expensive'
' i created command buttons to the cities, and each button adds to the flight price running total'
'Blake Bauer
'October 14th 2009'

Option Explicit
Dim City(1 To 10) As String, Price(1 To 600) As String, Ctr As Integer
Dim Answer As String


'ATL button with the flight price attached'
'it hides this frm and goes tothe ATL page'
Private Sub cmdAtlanta_Click()
'inputbox to make sure the people wanted to go here'
 Answer = InputBox("You want to travel to Atlanta? Y/N")
    If Answer = "Y" Then
        FlightPrice = 185
        frmFlights.Hide
    frmAtlanta.Show
    ElseIf Answer = "N" Then
        frmFlights.Show
        frmAtlanta.Hide
    End If
    
    
End Sub

'bahamas button'
'added the flight price to the flightprice totoal'
'hide and show frms as well'
Private Sub cmdBahamas_Click()
'inputbox to make sure the people wanted to go here'
 Answer = InputBox("You want to travel to Bahamas? Y/N")
    If Answer = "Y" Then
        FlightPrice = 500
         frmFlights.Hide
    frmBahamas.Show
     ElseIf Answer = "N" Then
        frmFlights.Show
        frmBahamas.Hide
    End If
    
   
End Sub
'Cancun Button'
'putting flight price on it'
Private Sub cmdCancun_Click()
'inputbox to make sure the people wanted to go here'
  Answer = InputBox("You want to travel to Cancun? Y/N")
    If Answer = "Y" Then
        FlightPrice = 450
        frmFlights.Hide
        frmCancun.Show
    ElseIf Answer = "N" Then
        frmFlights.Show
        frmCancun.Hide
    End If
    
    
End Sub

'quit Button'
Private Sub cmdEnd1_Click()
    End
End Sub

'LasVegas Button'
'flight Price and show and hide functions'
Private Sub cmdLasVegas_Click()
 Answer = InputBox("You want to travel to Las Vegas? Y/N")
    'inputbox to make sure the people wanted to go here'
    If Answer = "Y" Then
        FlightPrice = 200
        frmFlights.Hide
        frmLasVegas.Show
     ElseIf Answer = "N" Then
        frmFlights.Show
        frmLasVegas.Hide
    End If
   
End Sub
'Miami button'
'flight price and show and hide functions'
Private Sub cmdMiami_Click()
 Answer = InputBox("You want to travel to Miami? Y/N")
    'inputbox to make sure the people wanted to go here'
    If Answer = "Y" Then
        FlightPrice = 415
         frmFlights.Hide
            frmMiami.Show
     ElseIf Answer = "N" Then
        frmFlights.Show
        frmMiami.Hide
    End If

End Sub
'Panama City Button'
'flight price b eing added and showing and hidding frms'
Private Sub cmdPanamaCity_Click()
 Answer = InputBox("You want to travel to Panama City? Y/N")
    'inputbox to make sure the people wanted to go here'
    If Answer = "Y" Then
        FlightPrice = 395
         frmFlights.Hide
    frmPanamaCity.Show
     ElseIf Answer = "N" Then
        frmFlights.Show
        frmPanamaCity.Hide
    End If
   
End Sub


Private Sub cmdPrices_Click()
'declaring variabls so i can order the flights by price in a decending order'
Dim pass As Integer, pos As Integer, J As Integer
Dim tempCity As String, tempPrice As Single

Ctr = 0
'opening text file'
Open App.Path & "\Flights.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, City(Ctr), Price(Ctr)
Loop
Close #1
'decending order of flight prices'
For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Price(pos) < Price(pos + 1) Then
            tempCity = City(pos)
            City(pos) = City(pos + 1)
            City(pos + 1) = tempCity
            tempPrice = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = tempPrice
        End If
    Next pos
Next pass

'printing results'
picResults.Print "                                "
picResults.Print "City"; Tab(20); "Price of the Flight In Dollars"
picResults.Print "******************************************************************************"
'printing results'
For J = 1 To Ctr
             picResults.Print City(J); Tab(20); FormatCurrency(Price(J), 2)
    Next J

End Sub
'San Diego button'
'flight price and showing and hidding frms'
Private Sub cmdSanDiego_Click()

Answer = InputBox("You want to travel to San Diego? Y/N")
    If Answer = "Y" Then
        FlightPrice = 375
         frmFlights.Hide
    frmSanDiego.Show
     ElseIf Answer = "N" Then
        frmFlights.Show
        frmSanDiego.Hide
    End If
   

End Sub
