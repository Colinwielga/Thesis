VERSION 5.00
Begin VB.Form frmCheckout 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   17205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "End Your Magical Johnnie Travel Experience!!! :'("
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11880
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "See All The Info On the Fantastic Vacation That You Have Just Booked With Johnnie Travel!!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3960
      ScaleHeight     =   2835
      ScaleWidth      =   12675
      TabIndex        =   3
      Top             =   3360
      Width           =   12735
   End
   Begin VB.TextBox txtWholeName 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Johnnie Travel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1935
      Left            =   2280
      TabIndex        =   6
      Top             =   6960
      Width           =   11055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "First, Middle Initial, And Last Name ====>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Now It's Time For You To Check Out!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   9375
   End
End
Attribute VB_Name = "frmCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Vacation Planner
'Form Name: Destination
'Authors: Luke Gellerman and Tan Tuohy
'3/22/09
'This form checks the user out of the vacation planning program and gives them all their information that they had selected over the course
'of the entire program, with the costs of booking their hotel, arranging a flight, signing up for activities, and their total cost.
'All the variables that were declared Public are all pretty much used on this page, because there were running total that needed to be
'declared for all the pages

Option Explicit
Dim WholeName As String, n As Integer, CheckOutName As String       'declare global variables
Dim First As String, Middle As String, Last As String

Private Sub cmdCheckOut_Click()
    
    WholeName = txtWholeName.Text       'This code is a String Function which concantenates a person full name
    n = InStr(WholeName, " ")           'If a person typed in Luke S Gellerman, this program would display their name as Gellerman, L. S.
    First = Left(WholeName, n - 1)
    Last = Right(WholeName, Len(WholeName) - (n + 2))
    Middle = Mid(WholeName, n + 1, 1)
    CheckOutName = Last & ", " & Left(First, 1) & ". " & Left(Middle, 1) & "."
    
    CheckoutTotal = HotelTotal + FlightTotal + ActivitiesTotal
    
    'These few print statements tell the user exactly what they have signed up for through Johnnie Travel,
    'including all the costs of each individual portion, and then the final total cost.
    
    picResults.Print "Head of Party"; Tab(20); "Hotel Info And Cost"; Tab(60); "Flight Info And Cost"; Tab(112); "Activity Cost For Party"
    picResults.Print "*************************************************************************************************************************************************************************************************************"
    picResults.Print CheckOutName; Tab(20); FinalHotel; Tab(40); FormatCurrency(HotelTotal); Tab(60); "("; Flight; ")"; " "; Leaving; " "; FormatCurrency(FlightTotal); Tab(112); FormatCurrency(ActivitiesTotal)
    picResults.Print "                                                                                          "
    
    'printing out the final statement in the picture box saying a goodbye to the user
    picResults.Print "The total cost for your Johnnie Travel Vacation is"; " "; FormatCurrency(CheckoutTotal); "."
    picResults.Print "Thank you very much for choosing Johnnie Travel as your Vacation Planner!!!!!!!!!!!"
    picResults.Print "                                                                                          "
    
    'The ElseIf statement used here is dependent to what destination that the user chose earlier in the program,
    'because whichever location they chose, that location was assigned to the variable Location.
    'It will then print out a different line based on the destination that user chose.
    
    If Location = "St. Joe" Then
        picResults.Print "I guess you'll have some fun, but probably not that much in that janky town. I certainly wouldn't have picked "; Location; " as my travel destination, but whatever."
    ElseIf Location = "Badlands" Then
        picResults.Print "The "; Location; " are a much better choice than picking St. Joe."
    ElseIf Location = "Saskatchewan" Then
        picResults.Print Location; " Yea, you go to Canada and don't ever come back!!!"
    ElseIf Location = "Normandy" Then
        picResults.Print Location; " was by far the best vacation spot that we have to offer. You would have to be a sucker to go anywhere else."
    End If
    
End Sub

Private Sub cmdQuit_Click()
    End     'This code ends the program if they don't want to checkout
End Sub

Private Sub Form_Load()
    'This code centers the form on computer screen upon loading
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
End Sub
