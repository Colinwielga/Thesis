VERSION 5.00
Begin VB.Form frmBigBowl 
   BackColor       =   &H80000012&
   Caption         =   "Form9"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18765
   ForeColor       =   &H00000040&
   LinkTopic       =   "Form9"
   ScaleHeight     =   11730
   ScaleWidth      =   18765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMNLocal 
      BackColor       =   &H0080FFFF&
      Caption         =   "Minnesota Locations"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   11
      Text            =   "...a pioneering restaurant, vibrant in design and cusine"
      Top             =   1440
      Width           =   10335
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00404080&
      Caption         =   "Click here to place an order"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10800
      Width           =   3615
   End
   Begin VB.CommandButton cmdReservations 
      BackColor       =   &H00404080&
      Caption         =   "Make Reservations"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9240
      Width           =   3615
   End
   Begin VB.CommandButton cmdReturn4 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   13800
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit4 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   13800
      Width           =   1575
   End
   Begin VB.PictureBox picResults4 
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   4080
      ScaleHeight     =   10155
      ScaleWidth      =   12195
      TabIndex        =   6
      Top             =   3360
      Width           =   12255
   End
   Begin VB.CommandButton cmdLocations4 
      BackColor       =   &H00000080&
      Caption         =   "Click here if interested in finding other Big Bowl locations"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton cmdHistory4 
      BackColor       =   &H00000080&
      Caption         =   "Mission of Big Bowl"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdPrice4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Description of Popular Menu Items in Decending Order"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlpha4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Popular Menu Items Alphabetized"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdMenu4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Popular Menu Items"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Image imgBB3 
      Height          =   3015
      Left            =   0
      Picture         =   "Big Bowl.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image imgBB1 
      Height          =   3060
      Left            =   15840
      Picture         =   "Big Bowl.frx":5B32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblBigBowl 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Welcome to Big Bowl Resturant "
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   54.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3015
      Left            =   -840
      TabIndex        =   0
      Top             =   0
      Width           =   20775
   End
End
Attribute VB_Name = "frmBigBowl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CSCI VB Project: Big Bowl
'frmBigBowl
'Elizabeth K. Sturlaugson
'Due Date: Friday, March 28th, 2008
'The objective of this form is to introduce the various options that the user can perform.  The user can located popular menu items sort them by price and alphabetically.
    'The user can also make dinner reservations, order food on-line and search for other Big Bowl locations within the United States
'The form uses a variety of command buttons and message boxes that help rely information to the user.  Also, some commands allow the user to transport to different forms within the project.



Option Explicit
Dim NameArray(1 To 9) As String
Dim TypeArray(1 To 9) As String
Dim PriceArray(1 To 9) As String
Dim CTR As Integer


Private Sub cmdAlpha4_Click()
'sorts data alphabetically

'declare the variables
Dim correcttype As String
Dim correctname As String
Dim correctprice As String
Dim Pass As Integer
Dim Pos As Integer
Dim N As Integer

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
    If NameArray(Pos) > NameArray(Pos + 1) Then
    correctname = NameArray(Pos)
    NameArray(Pos) = NameArray(Pos + 1)
    NameArray(Pos + 1) = correctname
    
    correcttype = TypeArray(Pos)
    
    TypeArray(Pos) = TypeArray(Pos + 1)
    TypeArray(Pos + 1) = correcttype
    correctprice = PriceArray(Pos)
    PriceArray(Pos) = PriceArray(Pos + 1)
    PriceArray(Pos + 1) = correctprice
    
    End If
 Next Pos
 Next Pass
 
 
 picResults4.Cls
 
 For N = 1 To 9
 
 picResults4.Print NameArray(N); Tab(50); FormatCurrency(PriceArray(N))
 
 Next N



End Sub



Private Sub cmdHistory4_Click()
'current vision of Big Bowl

picResults4.Cls

picResults4.Print "Authentic Chinese and Thai flavors have always defined our menu, including"
picResults4.Print " using the finest soy sauces, sesame oils and cold-pressed peanut oil."
picResults4.Print " We have recently taken a giant step in our commitment to the enviroment"
picResults4.Print " and the local community."


End Sub

Private Sub cmdLocations4_Click()

'locates to see where other Big Bowl locations are in the United States
'user enters a state using an input box and the program tells is there is a location within that state

Dim State(1 To 3) As String
Dim NotState As String
Dim I As Integer
Dim CTR As Integer
Dim Found As Boolean

Open App.Path & "\BigBowlLocal.txt." For Input As #1
CTR = 1
Do Until EOF(1)
Input #1, State(CTR)
CTR = CTR + 1
Loop

NotState = InputBox("Please enter a state", "Locate a Big Bowl")
I = 0
Found = False

Do While ((Not Found) And (I < CTR - 1))
I = I + 1
If NotState = State(I) Then Found = True
Loop

If (Not Found) Then
MsgBox "Sorry, there isn't a Big Bowl in the state you selected.", , "Locations."
Else
    MsgBox "Yes, there is a Big Bowl in " & NotState, , "Locate a Big Bowl."
                    
                    
End If



Close #1





End Sub

Private Sub cmdMenu4_Click()

Dim N As Integer


'open the file for list of popular menu items
CTR = 0
Open App.Path & "/BigBowl.txt" For Input As #1
Do Until EOF(1)
CTR = CTR + 1

Input #1, NameArray(CTR), TypeArray(CTR), PriceArray(CTR)

Loop

Close #1

'creating the parallel arrays
For N = 1 To CTR

picResults4.Print NameArray(N); Tab(20); TypeArray(N); Tab(50); FormatCurrency(PriceArray(N))


Next N



End Sub

Private Sub cmdMNLocal_Click()
'displays message about other MN locations

MsgBox "There are 3 Big Bowl Locations in Minnesota: Edina, Minnetonka and Roseville.", , "Locations"



End Sub

Private Sub cmdOrder_Click()
'moves to another form

frmBigBowlOrder.Show
frmBigBowl.Hide




End Sub

Private Sub cmdPrice4_Click()
'sorts popular menu items by description in decending order

'declare the variables

Dim correcttype As String
Dim correctname As String
Dim correctprice As String
Dim Pass As Integer
Dim Pos As Integer
Dim N As Integer

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
    If TypeArray(Pos) < TypeArray(Pos + 1) Then
    correcttype = TypeArray(Pos)
    TypeArray(Pos) = TypeArray(Pos + 1)
    TypeArray(Pos + 1) = correcttype
    
    correctname = NameArray(Pos)
    
    NameArray(Pos) = NameArray(Pos + 1)
    NameArray(Pos + 1) = correctname
    correctprice = PriceArray(Pos)
    PriceArray(Pos) = PriceArray(Pos + 1)
    PriceArray(Pos + 1) = correctprice
    
    End If
 Next Pos
 Next Pass
 
 
 picResults4.Cls
 
 For N = 1 To 9
 
 picResults4.Print FormatCurrency(PriceArray(N)); Tab(10); NameArray(N); Tab(20); TypeArray(N)

 
 
 
 Next N
End Sub

Private Sub cmdQuit4_Click()
'quits
End
End Sub

Private Sub cmdReservations_Click()
'moves to another form
frmBigBowlReservations.Show
frmBigBowl.Hide

End Sub

Private Sub cmdReturn4_Click()
'moves to another form
frmTwinCities.Show
frmBigBowl.Hide
End Sub


