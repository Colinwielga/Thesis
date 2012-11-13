VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Hotel MainMenu"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCheckOut 
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4035
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   4200
      Width           =   6855
   End
   Begin VB.CommandButton cmdCheapest 
      Caption         =   "Sorting Prices"
      Height          =   1455
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdLoginPage 
      Caption         =   "Return to Log-in Page"
      Height          =   1455
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Guest Checkout"
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton cmdCheckin 
      Caption         =   "New Guest Checking In"
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: Hotel Checkin
'Form: Main Menu
'Authors: Ellen Jansen & Stuart Van Ess
'Date: March 28, 2008
'Purpose:   This is the main menu page. It is where the user can select what
'           they would like to do within the program. You can continue to the
'           new user page, continue to the sorting page, or check out a customer.

Option Explicit
Private Sub cmdCheapest_Click()
'shows the main menu and hides the sorting menu
    frmMainMenu.Hide
    frmsorting.Show
End Sub
Private Sub cmdCheckin_Click()
'shows the room size menu and hides the sorting menu
    frmRoomSize.Show
    frmMainMenu.Hide
End Sub
Private Sub cmdCheckout_Click()
'dimming all the variables needed within the form
    Dim CheckOutCTR As Integer
    Dim FirstNameArray(1 To 40) As String
    Dim LastNameArray(1 To 40) As String
    Dim Address1Array(1 To 40) As String
    Dim Address2array(1 To 40) As String
    Dim Address3array(1 To 40) As String
    Dim HomePhone1array(1 To 40) As String
    Dim HomePhone2array(1 To 40) As String
    Dim HomePhone3array(1 To 40) As String
    Dim CarArray(1 To 40) As String
    Dim LicensePlateArray(1 To 40) As String
    Dim NightsArray(1 To 40) As String
    Dim RoomArray(1 To 40) As String
    Dim Bill As Single
    Dim Tax As Single
    Dim Total As Single
    Dim FirstNameBox As String
    Dim LastNameBox As String
    Dim Found As Boolean
    Dim Pos As Integer

'sets found as false
    Found = False

'clears the picture box
    picCheckOut.Cls

'Asks the user what the guest's first name and last name is
'(the one who is trying to check out)
    FirstNameBox = InputBox("What is the guest's First Name?", "First Name", "")
    LastNameBox = InputBox("What is the guest's Last Name?", "LastName", "")

'opens the text file to be read
    Open App.Path & "\Guests.txt" For Input As #2

'reads the file into multiple different arrays
    CheckOutCTR = 0
    Do While Not EOF(2)
        CheckOutCTR = CheckOutCTR + 1
        Input #2, FirstNameArray(CheckOutCTR), LastNameArray(CheckOutCTR), Address1Array(CheckOutCTR), Address2array(CheckOutCTR), Address3array(CheckOutCTR), HomePhone1array(CheckOutCTR), HomePhone2array(CheckOutCTR), HomePhone3array(CheckOutCTR), CarArray(CheckOutCTR), LicensePlateArray(CheckOutCTR), NightsArray(CheckOutCTR), RoomArray(CheckOutCTR)
    Loop
    Close #2
    
'Compares the input name from above with the names in the arrays. If there is a
'match, the computer recognizes it and prints all the information tied with that
'name
    Do While (Not Found And Pos < CheckOutCTR)
        Pos = Pos + 1
        If FirstNameArray(Pos) = FirstNameBox And LastNameArray(Pos) = LastNameBox Then
            
            Found = True
            
            picCheckOut.Print "Bill for:"
            picCheckOut.Print
            picCheckOut.Print FirstNameArray(Pos); " "; LastNameArray(Pos)
            picCheckOut.Print
            picCheckOut.Print Address1Array(Pos)
            picCheckOut.Print Address2array(Pos)
            picCheckOut.Print Address3array(Pos)
            picCheckOut.Print
            picCheckOut.Print HomePhone1array(Pos); "-"; HomePhone2array(Pos); "-";
            picCheckOut.Print HomePhone3array(Pos)
            picCheckOut.Print
            picCheckOut.Print "Vehicle Make: "; CarArray(Pos)
            picCheckOut.Print "License Plate # "; LicensePlateArray(Pos)
            
    'Sets the prices for each of the rooms in order to be multiplied
            If RoomArray(Pos) = "Queen" Then
                RoomArray(Pos) = "139.99"
            ElseIf RoomArray(Pos) = "SmallSuite" Then
                RoomArray(Pos) = "179.99"
            ElseIf RoomArray(Pos) = "DoubleBed" Then
                RoomArray(Pos) = "119.99"
            ElseIf RoomArray(Pos) = "King" Then
                RoomArray(Pos) = "149.99"
            ElseIf RoomArray(Pos) = "MasterSuite" Then
                RoomArray(Pos) = "199.99"
            End If
            
    'Computes the customer's bill
            Bill = NightsArray(Pos) * RoomArray(Pos)
            Tax = 0.07 * Bill
            Total = Bill + Tax
            
            picCheckOut.Print
            
    'prints the customer's bill in the picture box
            picCheckOut.Print "Your Subtotal before tax equals "; FormatCurrency(Bill); "."
            picCheckOut.Print
            picCheckOut.Print "Tax equals "; FormatCurrency(Tax); "."
            picCheckOut.Print
            picCheckOut.Print "Your total is "; FormatCurrency(Total); "."
            
        Else: MsgBox "Please Enter a valid First Name or Last Name", , "Error"
        End If
    Loop
    
    
    
End Sub

Private Sub cmdLoginPage_Click()
'hides the main menu and shows the hotel menu
    frmMainMenu.Hide
    frmHotel.Show
End Sub
'When the form loads,
Private Sub Form_Load()
    picCheckOut.Cls
End Sub
