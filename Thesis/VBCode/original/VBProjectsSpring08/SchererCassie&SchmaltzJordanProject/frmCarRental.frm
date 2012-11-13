VERSION 5.00
Begin VB.Form frmCarRental 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   10905
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdforward 
      BackColor       =   &H0000FF00&
      Caption         =   "Continue to Next Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmdviper 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dodge Viper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8760
      Picture         =   "frmCarRental.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdshelby 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ford Shelby GT500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8760
      Picture         =   "frmCarRental.frx":1139
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdporsche 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Porsche 911 GT3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11400
      Picture         =   "frmCarRental.frx":1F85
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdsedona 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kia Sedona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8760
      Picture         =   "frmCarRental.frx":2983
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdaccord 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Honda Accord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6000
      Picture         =   "frmCarRental.frx":35FA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdfocus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ford Focus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3240
      Picture         =   "frmCarRental.frx":41BC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdeclipse 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mitsubishi Eclipse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   11400
      Picture         =   "frmCarRental.frx":4DAE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdcobalt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chevy Cobalt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      Picture         =   "frmCarRental.frx":5CE2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Car Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdsweepstakesNumbers 
      BackColor       =   &H0000FF00&
      Caption         =   "I'm Feelin' Lucky Sweepstakes"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   480
      ScaleHeight     =   2955
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   2640
      Width           =   7695
   End
   Begin VB.CommandButton cmdsweepstakes 
      BackColor       =   &H0000FF00&
      Caption         =   "Snazzy Name Sweepstakes!!!!"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label lbldirections 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Choose Your Car:"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   13
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Image Image8 
      Height          =   1590
      Left            =   720
      Picture         =   "frmCarRental.frx":694E
      Top             =   720
      Width           =   3600
   End
End
Attribute VB_Name = "frmCarRental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmCarRental
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/8/08
'Objective: This form allows the user to choose which kind of car they would like to rent.
'The user entered a sweepstakes to try to win a discount on their car rental. If they lost, they had the opportunity to enter a second sweepstakes.
'The Snazzy Name Sweepstakes allowed us to use a Match and Stop array search
'The I'm Feelin' Lucky Sweepstakes allowed us to use the numeric function INT()

Option Explicit

'Here we declared these variables globally

Dim winner As Boolean
Dim days As Integer
Dim discount As Single
Dim rate As Single

Private Sub cmdaccord_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 28.95
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub


Private Sub cmdcobalt_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 29.49
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdeclipse_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 38.95
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdfocus_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 24.49
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdforward_Click()

frmCarRental.Hide
frmEnd.Show

End Sub

Private Sub cmdload_Click()

'Here we declared our variables
'Inputed our car data and displayed it in the picture box

Dim make(1 To 10) As String
Dim model(1 To 10) As String
Dim cartype(1 To 10) As String
Dim price(1 To 10) As Single

Open App.Path & "\Cars.txt" For Input As #4

CTR = 0

Do While Not EOF(4)
    CTR = CTR + 1
    Input #4, make(CTR)
    Input #4, model(CTR)
    Input #4, cartype(CTR)
    Input #4, price(CTR)
Loop

picResults.Print "Make", "Model"; Tab(40); "Type"; Tab(60); "Price per day"
picResults.Print "*********************************************************************************************"

For J = 1 To CTR
    picResults.Print make(J), model(J); Tab(40); cartype(J); Tab(60); FormatCurrency(price(J))
Next J

'Only after the car data is loaded can the user select their vehicle choice

cmdviper.Visible = True
cmdsedona.Visible = True
cmdporsche.Visible = True
cmdshelby.Visible = True
cmdaccord.Visible = True
cmdeclipse.Visible = True
cmdfocus.Visible = True
cmdcobalt.Visible = True

End Sub

Private Sub cmdporsche_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 299.99
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdsedona_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 29.95
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdshelby_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 195.95
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub cmdsweepstakes_Click()

'Here we declared our variables

Dim snazzynames(1 To 60) As String
Dim name As String
Dim found As Boolean
winner = False

CTR = 0

'Here we inputed our names and put them into an array

Open App.Path & "\SnazzyNameSweepstakes.txt" For Input As #3

Do While Not EOF(3)
    CTR = CTR + 1
    Input #3, snazzynames(CTR)
Loop

'Here we used a match and stop search to search our array for the name inputed by the user
'If the name was found then the search stops and winner = true
'If the name was not found then they are not a winner, but another sweepstakes becomes available
'If they won, then they can load the car data

name = InputBox("Please type in your first name. If it is on our list you win!")
found = False
Pos = 0

Do While (Not found) And (Pos < CTR)
    Pos = Pos + 1
    If name = snazzynames(Pos) Then
        found = True
        MsgBox ("Congratulations you win 15% off your car rental!!")
        winner = True
        cmdload.Visible = True
    End If
Loop

If Not found Then
    MsgBox ("Sorry you are not a winner.")
    cmdsweepstakesNumbers.Visible = True
End If

cmdsweepstakes.Visible = False

End Sub

Private Sub cmdsweepstakesNumbers_Click()

'Here the user enters a number via inputbox
'if the numer is divisible by 17 then they win
'if they win, winner = true
'after they complete the sweepstakes they can then load the car data

winner = False
Dim number As Integer

number = InputBox("Please enter any number to see if you win the I'm Feelin' Lucky Sweepstakes!!")

If Int(number / 17) = number / 17 Then
    MsgBox (" Congratulations!! You Won 15% off your car rental!!")
    winner = True
Else: MsgBox ("Sorry you are not a winner.")
End If
    
cmdload.Visible = True
cmdsweepstakesNumbers.Visible = False
    
End Sub

Private Sub cmdviper_Click()

'Here we set our variables to the price
'to the number entered by the user via input box
'We used a boolean variable winner and an if then statement to determine what the discount is
'The total is displayed in a message box

rate = 235.9
days = InputBox("Pleas enter the number of days you wish to rent the vehicle.")

If winner = True Then
    discount = 0.85
Else: discount = 1
End If

carRentaltotal = rate * days * discount

MsgBox ("Your total is " & FormatCurrency(carRentaltotal) & ".")

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub


