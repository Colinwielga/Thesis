VERSION 5.00
Begin VB.Form frmPrices 
   BackColor       =   &H000080FF&
   Caption         =   "Compare Deals and Prices"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMonth 
      Caption         =   "Search for special offers by month of the year"
      Height          =   975
      Left            =   5640
      TabIndex        =   8
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmdOneway 
      Caption         =   "Search one-way ticket prices by cities"
      Height          =   975
      Left            =   7560
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtboxFinish 
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtboxStart 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdDisplayPrices 
      Caption         =   "Search round-trip ticket prices by city"
      Height          =   975
      Left            =   7560
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picOutbox 
      AutoRedraw      =   -1  'True
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6555
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdMainMenu2 
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   7080
      TabIndex        =   0
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label lblMembername 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H000080FF&
      Caption         =   "**All prices displayed are calculated without tax!"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblName 
      BackColor       =   &H000080FF&
      Caption         =   "By: Chaise VanOverbeke"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   4800
      Picture         =   "frmPrices.frx":0000
      Top             =   5280
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Look here for special offers and packages of the month..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "Enter destination point (city):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Enter starting point (city):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Airline Option(Project1.vbp)
'Form Name : frmPrices(frmPrices.frm)
'Author: Chaise VanOverbeke
'Date : Thursday October 27, 2005
'Purpose of the form:  This form allows the user to see the available flights that Chaise Van
                    'Air offers and then look up their corresponding prices, both round-trip
                    'and one-way, if they exist.  This form also allows the user to search by
                    'months to see if there are any deals that are relevant to their particular
                    'flight.

Option Explicit
    Dim Start(1 To 50) As String
    Dim Finish(1 To 50) As String
    Dim Price(1 To 50) As Single
    Dim CTR As Integer

Private Sub cmdDisplayPrices_Click()
    Dim notFound As Boolean
    Dim I As Integer
    Dim S As String
    Dim F As String
    S = txtboxStart.Text
    F = txtboxFinish.Text
    I = 0
    notFound = True   'sets notfound = true
    Do While notFound And I < 50   'searches the array and keeps track with the number(I)
        I = I + 1
        If S = Start(I) And F = Finish(I) Then notFound = False     'user enters a starting point and destination that Chaise Van Air does offer'
    Loop
    If notFound Then    'notfound = true, the flight the user inputed does not exist.
        MsgBox "We do not offer that flight.", , "No Flights"
    Else    'If the other qualification is not met, then the flight was found, and a message is displayed in the following messagebox.
        MsgBox "The price round trip is " & Price(I) * 2 & ".", , "Found Price"    'a message box is displayed that takes the price of the one-way ticket and multiplies it by two, to get the roundtrip price of the flight the user selected.
    End If

        
        
    
    
End Sub

Private Sub cmdMainMenu2_Click()
    frmPrices.Hide
    frmMainForm1.Show
End Sub

Private Sub cmdMonth_Click()
    Dim Months As String
    Dim January As String   'the user enters a number into the case statement and it compares all the cases until a match is found and then a messagebox is displayed for the corresponding month.
    Months = InputBox("Input the number corresponding to the month you are planning to take your trip in, for ex. January = 1, May = 5")
    Select Case Months
    Case 1
        MsgBox "All flights to California are 20% off", , "Great Deals"
    Case 2
        MsgBox "All Roundtrip flights to anywhere in the United States are 15% off", , "Great Deals"
    Case 3
        MsgBox "All flights to Massachusetts and Maine are 20% off", , "Great Deals"
    Case 4
        MsgBox "All one-way tickets to anywhere in the U.S are 18% off", , "Great Deals"
    Case 5
        MsgBox "All flights to Minnesota and Wisconsin are 25% off", , "Great Deals"
    Case 6
        MsgBox "All flights to Utah and Nevada are 10% off", , "Great Deals"
    Case 7
        MsgBox "All flights to Virginia and Georgia are 20% off", , "Great Deals"
    Case 8
        MsgBox "All Roundtrip flights to anywhere in the United States are 15% off", , "Great Deals"
    Case 9
        MsgBox "All flights to Texas are 23% off", , "Great Deals"
    Case 10
        MsgBox "All flights to Arizona are 20% off", , "Great Deals"
    Case 11
        MsgBox "All one-way tickets to anywhere in the U.S are 18% off", , "Great Deals"
    Case 12
        MsgBox "There are no deals during the month of December", , "Great Deals"
    Case Else   'the user inputed a number that does not exist within the correct range.
       MsgBox "You have entered a number that doesn't correspond with a month!", , "Error"
    End Select
End Sub

Private Sub cmdOneway_Click()
 Dim notFound As Boolean
    Dim I As Integer
    Dim S As String
    Dim F As String
    S = txtboxStart.Text
    F = txtboxFinish.Text
    I = 0
    notFound = True     'sets notfound = true
    Do While notFound And I < 50   'searches the array and keeps track with the number(I)
        I = I + 1
        If S = Start(I) And F = Finish(I) Then notFound = False    'user enters a starting point and destination that Chaise Van Air does offer.
    Loop
    If notFound Then
        MsgBox "We do not offer that flight.", , "No Flights"   'notfound = true, the flight the user inputed does not exist.
    Else
        MsgBox "The price one-way is " & Price(I) & ".", , "Found Price"   'a message box is displayed with the corresponding price of the one-way ticket
    End If

End Sub

Private Sub Form_Load()
Open App.Path & "\Trips.txt" For Input As #1   'opens the file of all the flights Chaise Van Air offers
CTR = 0
picOutbox.Print "Start", "Finish", "Start", "Finish"
picOutbox.Print "************************************************************************************"
    Do Until EOF(1)    'Cycles through the data until the end of the file
        CTR = CTR + 1
        Input #1, Start(CTR), Finish(CTR), Price(CTR)
        picOutbox.Print Start(CTR), Finish(CTR); ", ",  'prints in a picture box all of the starting points and destinations of the flights that Chaise Van Air offers.
        If CTR Mod 2 = 0 Then
            picOutbox.Print
        End If
    Loop
Close #1   'close file when done reading the array.

End Sub
