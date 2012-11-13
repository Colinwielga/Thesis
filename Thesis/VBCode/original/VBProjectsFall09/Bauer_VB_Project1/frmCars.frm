VERSION 5.00
Begin VB.Form frmCars 
   BackColor       =   &H0080FF80&
   Caption         =   "Final Bill"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Quit"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search For a Color of a Car"
      BeginProperty Font 
         Name            =   "Doulos SIL"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdAlpabetical 
      BackColor       =   &H000000FF&
      Caption         =   "Cars listed Alphabetically"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrice 
      BackColor       =   &H000080FF&
      Caption         =   "Rental Cars listed by Price"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrius 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Toyota Prius"
      Height          =   1215
      Left            =   10200
      Picture         =   "frmCars.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCadiliac 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cadillac Escalade EXT"
      Height          =   1215
      Left            =   10200
      Picture         =   "frmCars.frx":0A23
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdBronco 
      BackColor       =   &H000000FF&
      Caption         =   "Ford Bronco"
      Height          =   1215
      Left            =   10200
      Picture         =   "frmCars.frx":1317
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdAudi 
      BackColor       =   &H00FFFF80&
      Caption         =   "Audi R8"
      Height          =   1215
      Left            =   360
      Picture         =   "frmCars.frx":1D29
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdSubaru 
      BackColor       =   &H00404040&
      Caption         =   "Subaru STI"
      Height          =   1215
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmCars.frx":26CD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmPorsche 
      BackColor       =   &H8000000A&
      Caption         =   "Porsche 911"
      Height          =   1215
      Left            =   10200
      Picture         =   "frmCars.frx":31DB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdFord 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ford Mustang"
      Height          =   1215
      Left            =   10200
      Picture         =   "frmCars.frx":375D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   6975
      Left            =   3000
      ScaleHeight     =   6915
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
   End
   Begin VB.CommandButton cmdOld 
      BackColor       =   &H00808080&
      Caption         =   "Final Bill"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label lblRental 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rental Cars"
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   9
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rental Car Page'
'This page is all about car rentals'
'This is where i did evertyhing from searching to listing to printing'
'October 18th 2009
'BLake Bauer'
'Declaring variables that will be used through out this frm'
Option Explicit
Dim Car(1 To 10) As String, Price(1 To 10) As Integer, Color(1 To 10) As String, Ctr As Integer
Dim CarDay As Integer, Yes As String
Private Sub cmdAlpabetical_Click()
'Declaring variabls for a alpabetical ordering'
Dim pass As Integer, pos As Integer, J As Integer
Dim tempCar As String, tempPrice As Single

picResults.Cls


Ctr = 0
'opening the text file'
Open App.Path & "\Cars.txt" For Input As #1
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Car(Ctr), Price(Ctr), Color(Ctr)
Loop
Close #1

'ordering the cars alpabetically'

For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Car(pos) > Car(pos + 1) Then
            tempPrice = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = tempPrice
            tempCar = Car(pos)
            Car(pos) = Car(pos + 1)
            Car(pos + 1) = tempCar
        End If
    Next pos
Next pass

'printing my results'
picResults.Print "                                "
picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"
picResults.Print "******************************************************************************"
'printing results'
For J = 1 To Ctr
             picResults.Print Car(J); Tab(20); FormatCurrency(Price(J), 2)
    Next J

End Sub
'this displays the info about the audi'
Private Sub cmdAudi_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Audi R8"; Tab(20); 400#; , , ; , , ; "Black"
    Yes = InputBox("Would you Like To Rent The Audi R8? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Audi R8?")
        Else
        End If
    CarPrice = CarDay * 400
End Sub
'this displays the info about the Bronco'
Private Sub cmdBronco_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Ford Bronco"; Tab(20); 150#; , , ; , , ; "Red"
    Yes = InputBox("Would you Like To Rent The Ford Bronco? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Ford Bronco?")
        Else
        End If
        CarPrice = CarDay * 150
End Sub
'this displays the info about the cadillac'
Private Sub cmdCadiliac_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Cadillac EXT"; Tab(20); 260#; , , ; , , ; "Blue"
    Yes = InputBox("Would you Like To Rent The Cadillac EXT? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Cadillac EXT?")
        Else
        End If
        CarPrice = CarDay * 260
End Sub
'decalring my variables for a searching task'
Private Sub cmdColor_Click()
Dim Found As Boolean
Dim Ctr As Integer, tempColor As String
Dim I As Integer
'clearing the pic box'
picResults.Cls



Ctr = 0
'opeing text file'
Open App.Path & "\Cars.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Car(Ctr), Price(Ctr), Color(Ctr)
Loop
Close #1

'creating a inputbox'
tempColor = InputBox("What color of a car would you like? (Black, Bue, Red, White, Green)(!Capitilize the first letter!)", "COLOR OF CAR")

picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
picResults.Print "****************************************************************************************************"
'searching for colors in a text file'
For I = 1 To Ctr 'searing the enire list'
        If tempColor = Color(I) Then
            picResults.Print Car(I); Tab(20); FormatCurrency(Price(I), 2); , , ; , , ; Color(I)
        End If
    
Next I



End Sub

'quit button'
Private Sub cmdEnd_Click()
    End
End Sub
'this displays the info about the Ford Mustang'
Private Sub cmdFord_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "For Mustang"; Tab(20); 250#; , , ; , , ; "Black"
    Yes = InputBox("Would you Like To Rent The Mustang? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Mustang?")
        Else
        End If
        CarPrice = CarDay * 250
End Sub
'going to the last frm'
Private Sub cmdOld_Click()
    frmCars.Hide
    frmEnd.Show
    
End Sub

'I am declaring my variables for a list of ordering by price'

Private Sub cmdPrice_Click()
Dim pass As Integer, pos As Integer, J As Integer
Dim tempCar As String, tempPrice As Single

picResults.Cls



Ctr = 0
'opening the text file agaim'
Open App.Path & "\Cars.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Car(Ctr), Price(Ctr), Color(Ctr)
Loop
Close #1
'ordering the cars in decending price orders'
For pass = 1 To Ctr - 1
    For pos = 1 To Ctr - pass
        If Price(pos) < Price(pos + 1) Then
            tempCar = Car(pos)
            Car(pos) = Car(pos + 1)
            Car(pos + 1) = tempCar
            tempPrice = Price(pos)
            Price(pos) = Price(pos + 1)
            Price(pos + 1) = tempPrice
        End If
    Next pos
Next pass

'printing results'
picResults.Print "                                "
picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"
picResults.Print "******************************************************************************"
'printing results'
For J = 1 To Ctr
             picResults.Print Car(J); Tab(20); FormatCurrency(Price(J), 2)
    Next J


End Sub
'this displays the info about the prius'
Private Sub cmdPrius_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Tyota Prius"; Tab(20); 200#; , , ; , , ; "Green"
    Yes = InputBox("Would you Like To Rent The Toyota Prius? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Toyota Prius?")
        Else
        End If
        CarPrice = CarDay * 200
End Sub
'this displays the info about the subaru'
Private Sub cmdSubaru_Click()
  picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Subaru STI"; Tab(20); 275#; , , ; , , ; "White"
    Yes = InputBox("Would you Like To Rent The Subaru STI? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Subaru STI?")
        Else
        End If
    CarPrice = CarDay * 275
    
End Sub
'this displays the info about the porsche'
Private Sub cmPorsche_Click()
picResults.Cls
    picResults.Print "Car"; Tab(20); "Price Per Day of Car Rental"; , , ; "Color"
    picResults.Print "****************************************************************************************************"
    picResults.Print "Porsche 911"; Tab(20); 350#; , , ; , , ; "Red"
    Yes = InputBox("Would you Like To Rent The Porsche 911? (Y/N)")
        If Yes = "Y" Then
            CarDay = InputBox("How Many Days Would You like to Rent the Porshce 911?")
        Else
        End If
        CarPrice = CarDay * 350
End Sub
