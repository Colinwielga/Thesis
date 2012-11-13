VERSION 5.00
Begin VB.Form frmHotel 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   ScaleHeight     =   10860
   ScaleWidth      =   15480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "If You Want to Go Back and Change Your Travel Destination, Do It Now!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6480
      TabIndex        =   17
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton cmdNextActivityOptionPage 
      Caption         =   "Let's Get You Hooked Up With a Flight!!!!"
      Height          =   1455
      Left            =   12840
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdMatchStopSearch 
      Caption         =   "Enter the Number Of the Hotel That You Would Like To Stay At!!!!"
      Height          =   1335
      Left            =   10320
      TabIndex        =   10
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAscend 
      Caption         =   "See All the Prices From Least Expensive To Most Expensive =========>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Wondeful Johnnie Travel Experience :'("
      Height          =   1335
      Left            =   13320
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdHotels 
      Caption         =   "Check out all our wonderful hotels that Johnnie Travel has to offer!!!! ======>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox cboPeople 
      Height          =   315
      ItemData        =   "Hotel.frx":0000
      Left            =   5520
      List            =   "Hotel.frx":001C
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtNights 
      BackColor       =   &H000000C0&
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   5160
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0E0FF&
      Height          =   2895
      Left            =   3000
      ScaleHeight     =   2835
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   16
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF00FF&
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   4365
      Left            =   9000
      Picture         =   "Hotel.frx":0038
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   12
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   11
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   600
      Picture         =   "Hotel.frx":18571
      Top             =   6120
      Width           =   5715
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   $"Hotel.frx":306E7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Lets See What Hotel You Want To Stay At!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   10335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "How many nights will your party be staying?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "How many people will be staying in your hotel room?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   4320
      Width           =   3375
   End
End
Attribute VB_Name = "frmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/21/09
'The user selects how many people will be staying in the hotel and the number of nights
'The list of hotels, their cost per person for a single night, and the hotel number is loaded into three arrays
'The array is sorted in ascending order according to cost per person for a single night
'The user selects a hotel to stay at based on the hotel number through an InputBox
'The match and stop search then searches for the matching hotel ID number that was loaded into an array.
'The program calculates the total booking cost for the hotel based on the number of people multiplied by the number of
'nights multiplied by the cost of the hotel. There is a discount of 15% for the total booking fee if it is over $500 dollars.


Option Explicit
Dim Nights As Integer, People As Integer            'declare all the global variables
Dim Hotels(1 To 50) As String, Cost(1 To 50) As Integer, Num(1 To 50) As Integer
Dim Temp As Integer, Temp2 As String, Temp1 As Integer, HotelPrice As Single, HotelPriceD As Single
Dim HotelNum As Integer, Pass As Integer, Pos As Integer


Private Sub cmdAscend_Click()
    
    picResults.Cls      'clears the picture box of all data
    picResults.Print " Number of Hotel "; Tab(20); " Choices for Hotels ", "Cost Per Person For a Single Night"     'prints both lines when
    picResults.Print "****************************************************************************************"     'button is first pressed
    
    For Pass = 1 To CTR - 1     'keeps track of how many passes
        For Pos = 1 To CTR - Pass       'keeps track of how many comparisons
            If Cost(Pos) > Cost(Pos + 1) Then
                Temp = Num(Pos)     'exchanges values if out of order
                Num(Pos) = Num(Pos + 1)
                Num(Pos + 1) = Temp
                Temp1 = Cost(Pos)       'exchanges values if out of order
                Cost(Pos) = Cost(Pos + 1)
                Cost(Pos + 1) = Temp1
                Temp2 = Hotels(Pos)     'exchanges values if out of order
                Hotels(Pos) = Hotels(Pos + 1)
                Hotels(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
                
    For I = 1 To CTR
        picResults.Print Num(I); Tab(20); Hotels(I), Tab(50); FormatCurrency(Cost(I))    'prints all of the data, but in ascending order
    Next I
            
End Sub

Private Sub cmdBack_Click()
    'hides current form and allows the user to go back and select a new travel destination in the Destination form
    
    frmHotel.Hide
    frmDestination.Show
    
End Sub

Private Sub cmdMatchStopSearch_Click()
    'This is a match and stop search used to look throughout the entire list and find the number that the user enters,
    'which corresponds with the hotel number
    'Based on the hotel number, the cost per person for a single night is now assigned to "Cost(I)"
    'and the calculation for the total booking cost for the hotel can be taken care of and
    'is then displayed to the user through a message box

    HotelNum = InputBox("Enter the number of the hotel you wish to stay at!!", "Hotel Selection")
    I = 0
    Found = False
    
    Nights = txtNights.Text     'Nights and People are found through a text box and combo box
    People = cboPeople.Text
    
    Do While ((Not Found) And (I < CTR))        'match and stop searches the list for the Hotel Number that was entered by the user
        I = I + 1                               'so that it will match up the cost of that hotel, so that they can choose their hotel
        If HotelNum = Num(I) Then               'by typing in one number, instead of typing in the whole name of the hotel
            Found = True
        End If
    Loop
    
    FinalHotel = Hotels(I)      'Final Hotel needed on the last checkout form so it is public
    HotelPrice = Cost(I) * Nights * People      'declaring what HotelPrice is equal to
    HotelPriceD = (Cost(I) * Nights * People) - ((Cost(I) * Nights * People) * 0.15)        'The price of the hotel with the 15% discount
    
    'If boolean expression is still false, then one message box comes up, otherwise found = true and then a different message box comes up
    'prompting the user the info with booking their hotel
    
    If (Not Found) Then
        MsgBox HotelNum & " is not a number of one of our hotels!!! Try Again!!"
    Else
        MsgBox "Thank you for choosing " & Hotels(I) & " as your hotel choice. The cost per person for a single night is " & FormatCurrency(Cost(I)) & "."
            If HotelPrice > 500 Then
                MsgBox "The total booking fee for your party to stay at " & Hotels(I) & " is " & FormatCurrency(HotelPriceD) & "."
                HotelTotal = HotelPriceD
                CheckoutTotal = CheckoutTotal + HotelPriceD
            Else
                MsgBox "The total booking fee for your party to stay at " & Hotels(I) & " is " & FormatCurrency(HotelPrice) & "."
                HotelTotal = HotelPrice
                CheckoutTotal = CheckoutTotal + HotelTotal
            End If
    End If
    
    'public HotelTotal used in last form, and a running tally of CheckoutTotal
    
End Sub

Private Sub cmdHotels_Click()

    Open App.Path & "\Hotels.txt" For Input As #1
    
    CTR = 0             'opens the file and reads it into three arrays and then closes the file
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, Num(CTR), Hotels(CTR), Cost(CTR)
    Loop
    
    Close #1        'closes file once it has been put into the arrays
    
    'prints these headings before all the data from the file is printed
    picResults.Print " Number of Hotel "; Tab(20); " Choices for Hotels ", "Cost Per Person For a Single Night"
    picResults.Print "****************************************************************************************"
    
    'prints the hotel number, names of the hotels, and thecost of staying at the hotel from the file
    For I = 1 To CTR
        picResults.Print Num(I); Tab(20); Hotels(I), Tab(50); FormatCurrency(Cost(I))
    Next I


End Sub

Private Sub cmdNextActivityOptionPage_Click()
    'hides current form and shows flight form
    
   frmHotel.Hide
   frmFlight.Show
    
End Sub

Private Sub cmdQuit_Click()
    End     'ends the entire program when the user clicks the button
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    
End Sub
