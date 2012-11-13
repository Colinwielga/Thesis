VERSION 5.00
Begin VB.Form frmBestCar 
   BackColor       =   &H0080C0FF&
   Caption         =   "Kim Nguyen"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H0080FF80&
      Caption         =   "Go Back To Homepage"
      Height          =   615
      Left            =   4800
      MaskColor       =   &H0080FF80&
      TabIndex        =   14
      Top             =   7440
      Width           =   1935
   End
   Begin VB.PictureBox picChosenCar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   4680
      ScaleHeight     =   6675
      ScaleWidth      =   5715
      TabIndex        =   13
      Top             =   600
      Width           =   5775
   End
   Begin VB.PictureBox Picture10 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmBestCar.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   12
      Top             =   6960
      Width           =   2055
   End
   Begin VB.PictureBox Picture9 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmBestCar.frx":7062
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   11
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture8 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmBestCar.frx":E0C4
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   10
      Top             =   3600
      Width           =   2055
   End
   Begin VB.PictureBox Picture7 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmBestCar.frx":15126
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox Picture6 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmBestCar.frx":1C188
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   0
      Picture         =   "frmBestCar.frx":231EA
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   0
      Picture         =   "frmBestCar.frx":2BF8C
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   3600
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   0
      Picture         =   "frmBestCar.frx":32FEE
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdBestCar 
      Caption         =   "Which 3 Cars Do You Think an Enthusiast Would Buy"
      Height          =   615
      Left            =   8160
      TabIndex        =   4
      Top             =   7440
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      Picture         =   "frmBestCar.frx":3A050
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   1095
      Left            =   0
      Picture         =   "frmBestCar.frx":410B2
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7080
      TabIndex        =   1
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdAccending 
      Caption         =   "List the Price of The Car In Ascending Order"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pick 3 Cars That You  Think an Enthusiast Would Buy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Caption         =   "10"
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Caption         =   "9"
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   "8"
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "7"
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "6"
      Height          =   255
      Left            =   2760
      TabIndex        =   30
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "5"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "4"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "3"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label pic2 
      BackColor       =   &H0080C0FF&
      Caption         =   "2"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label pic1 
      BackColor       =   &H0080C0FF&
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "BMW 3 Series"
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mercedes-Benz"
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      Caption         =   "Chrysler Prowler"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080C0FF&
      Caption         =   "Audi A4"
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080C0FF&
      Caption         =   "BMW M3"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080C0FF&
      Caption         =   "BMW 5-Series"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080C0FF&
      Caption         =   "Volvo C70"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080C0FF&
      Caption         =   "BMW Z4"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mercedes-Benz SLK"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080C0FF&
      Caption         =   "Audi A6"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   8160
      Width           =   1335
   End
End
Attribute VB_Name = "frmBestCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Online Car Dealer
'Form Name : frmBestCar (BestDeal1.frm)
'Author: Kim Nguyen
'Date Written: October 29, 2003
'Purpose of Form: To let the customer to view all of the cars available on one page
                'Let the user to input the infomartion by using inputbox to see which are that an Enuthasist would buy
                'then the output will print out the preferable cars that an Enuthansist would buy
                'This form also let the user to sort the price of the car from lowest to hightest price
                'Also there are Quit button which to end the program, and botton that will bring the user back to the homepage
                

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim Order(1 To 10) As Integer
Dim Year(1 To 10) As Integer
Dim CarName(1 To 10) As String
Dim Price(1 To 10) As Long
Dim Class(1 To 10) As String
Dim Cylinders(1 To 10) As String
Dim Horsepower(1 To 10) As Single
Dim StandardTransmission(1 To 10) As String
Dim Drivetrain(1 To 10) As String
Dim PicName(1 To 10) As String
Dim I As Integer


'sort the price of the cars in an array from lowest ot highest using bubble sort

Private Sub cmdAccending_Click()
picChosenCar.Cls                    'clear the picture box
'Declare all the variables before use
Dim temp As Double
Dim tempstr As String
Dim comp As Double
Dim Pass As Integer
picChosenCar.Print " "; "Name"; Tab(23), "Price"
picChosenCar.Print

For Pass = 1 To 10  'number of passes through the list
    For comp = 1 To 10 - Pass   'Number of comparisons for each pass.
        If Price(comp) > Price(comp + 1) Then   'compare adjacent price
            temp = Price(comp)  'swap if necessary
            Price(comp) = Price(comp + 1)
            Price(comp + 1) = temp
            tempstr = CarName(comp)
            CarName(comp) = CarName(comp + 1)
            CarName(comp + 1) = tempstr
            
        End If
    Next comp
Next Pass

'Display the sorted list of CarName and Price
'in the picture box named "picChosenCar".

For I = 1 To 10

    picChosenCar.Print (I); CarName(I); Tab(23), FormatCurrency(Price(I), 2)
    picChosenCar.Print
Next I
Close #1
End Sub

'Open the file and put it into an array
Public Sub Form_Load()
Open Path & "Cars.txt" For Input As #2
For I = 1 To 10
    Input #2, Year(I), CarName(I), Price(I), Class(I), Cylinders(I), Horsepower(I), StandardTransmission(I), Drivetrain(I), PicName(I)
Next I
Close #2
End Sub
'Ask the user to input 3 (number of the car) that they think an Enthusaist would buy.
'The numbers are used to sort the price to go alone with it.
'The names  and the price are also printed in the
'picture box right after the user types them in the input box.
'Then the list of car that Enthusiast will buy will be printed also

Private Sub cmdBestCar_Click()
picChosenCar.Cls
I = InputBox("Enter The Number of The 1st Car That You Think an Enthusiast Would Buy")
picChosenCar.Print "Name of The Car That You Picked"; Tab(50); "Price"
picChosenCar.Print "____________________________________________________"
picChosenCar.Print "1."; CarName(I); Tab(50); FormatCurrency(Price(I))
picChosenCar.Print
I = InputBox("The Number of The 2nd Car That You Think an Enthusiast Would Buy ")
picChosenCar.Print "2."; CarName(I); Tab(50); FormatCurrency(Price(I))
picChosenCar.Print
I = InputBox("The Number of the 3rd Car That You Think an Enthusiast Would Buy")
picChosenCar.Print "3."; CarName(I); Tab(50); FormatCurrency(Price(I))
picChosenCar.Print "------------------------------------------------------------------------------------------------------"
picChosenCar.Print
picChosenCar.Print "Enthusiast like to pick cars base on"
picChosenCar.Print "* trades comfort and dependability for excitement"
picChosenCar.Print "* prefers powerful, well-engineered cars"
picChosenCar.Print
picChosenCar.Print "Name of The Car That An Enthusiast Picked"; Tab(50); "Price"
picChosenCar.Print
picChosenCar.Print "1."; CarName(5); Tab(50); FormatCurrency(Price(5))
picChosenCar.Print
picChosenCar.Print "2."; CarName(7); Tab(50); FormatCurrency(Price(7))
picChosenCar.Print
picChosenCar.Print "3."; CarName(9); Tab(50); FormatCurrency(Price(9))

End Sub
'hide the Info form and show the BestCar screen
Private Sub cmdHome_Click()
frmInfo.Show
frmBestCar.Hide
End Sub
'End the program now
Private Sub cmdQuit_Click()
End
End Sub



