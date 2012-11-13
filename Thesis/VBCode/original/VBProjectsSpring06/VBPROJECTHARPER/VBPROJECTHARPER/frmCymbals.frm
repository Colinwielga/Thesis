VERSION 5.00
Begin VB.Form frmCymbals 
   BackColor       =   &H80000012&
   Caption         =   "Buy Cymbals"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Home Page"
      Height          =   1335
      Left            =   480
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Total To Cart"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7560
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Total"
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdTotalCym 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cymbal Total"
      Height          =   615
      Left            =   6840
      TabIndex        =   9
      Top             =   5760
      Width           =   1815
   End
   Begin VB.PictureBox picResultscymbals 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   6840
      ScaleHeight     =   4755
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label9 
      Caption         =   "By: Ben Harper"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   780
      Left            =   2760
      Picture         =   "frmCymbals.frx":0000
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image Z8 
      Height          =   1320
      Left            =   4560
      Picture         =   "frmCymbals.frx":4962
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Image Z7 
      Height          =   1275
      Left            =   4560
      Picture         =   "frmCymbals.frx":C564
      Top             =   3840
      Width           =   1800
   End
   Begin VB.Image Z6 
      Height          =   1335
      Left            =   4560
      Picture         =   "frmCymbals.frx":13D2E
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Image Z5 
      Height          =   1410
      Left            =   4560
      Picture         =   "frmCymbals.frx":1BA98
      Top             =   720
      Width           =   1800
   End
   Begin VB.Image Z4 
      Height          =   1470
      Left            =   1080
      Picture         =   "frmCymbals.frx":23F0A
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Image Z3 
      Height          =   1245
      Left            =   1080
      Picture         =   "frmCymbals.frx":2C91C
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Image Z2 
      Height          =   1290
      Left            =   1080
      Picture         =   "frmCymbals.frx":33E16
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Image Z1 
      Height          =   1335
      Left            =   1080
      Picture         =   "frmCymbals.frx":3B748
      Top             =   720
      Width           =   1800
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian A series Rides"
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian K Custom Crash $179.80"
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian Oriental China $249.49"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian A Custom Splash $129"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian Custom K Hi-Hats $212.99"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian ZBT Crash Cymbals $199.50"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian Custom A Hi-Hats $159.99"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zildjian Custom A Crashes"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmCymbals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClear_Click()  ' clear and set all cymbal values to zero
    picResultscymbals.Cls
    Cymbalsum = 0
    cmdSend.Visible = False
    cmdTotalCym.Visible = True
End Sub

Private Sub cmdReturn_Click() 'return to home page
    frmHomePage.Visible = True
    frmCymbals.Visible = False
End Sub

Private Sub cmdSend_Click() 'sends cymbals and total price to cart
    frmCart.Visible = True
    frmCymbals.Visible = False
    frmCart.picResults.Print "Cymbal Purchases", "Total"
    frmCart.picResults.Print "*************************************************"
    frmCart.picResults.Print "Your total for Cymbals is: ", FormatCurrency(Cymbalsum)
End Sub





Private Sub cmdTotalCym_Click()      'adds cymbal to cymbal total
    picResultscymbals.Print "*************************************"
    picResultscymbals.Print " Your total Cymbal cost is: ", FormatCurrency(Cymbalsum)
    cmdTotalCym.Visible = False
    cmdSend.Visible = True
End Sub

Private Sub Z1_Click()
    Dim InchesArray(1 To 25) As Integer
    Dim NamesArray(1 To 12) As String
    Dim PricesArray(1 To 300) As Single
    Dim Found As Boolean
    Dim Pos, Search As Integer
        
    Open App.Path & "\ZildjianACrash.txt" For Input As #1 'opens cymbal file and places into three arrays
        Pos = 0
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, InchesArray(Pos), NamesArray(Pos), PricesArray(Pos)
    Loop
    Close #1
    
    Found = False
    Pos = 0
    Search = InputBox("enter the size (17-19 inches) of cymbal you would like", "Cymbal Size")
        
    Do While Found = False
        Pos = Pos + 1
        If Search = InchesArray(Pos) Then
        Found = True
        End If
    Loop
    If Found = True Then
        picResultscymbals.Print NamesArray(Pos), FormatCurrency(PricesArray(Pos))
        Cymbalsum = Cymbalsum + PricesArray(Pos)
    End If
End Sub

Private Sub Z2_Click()   'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian Custom A Hi-Hat", FormatCurrency(159.99, 2)
    Cymbalsum = Cymbalsum + 159.99
End Sub

Private Sub Z3_Click()   'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian ZBT Crash Cymbals", FormatCurrency(199.5, 2)
    Cymbalsum = Cymbalsum + 199.5
End Sub

Private Sub Z4_Click()    'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian A Custom Splash", FormatCurrency(129, 2)
    Cymbalsum = Cymbalsum + 129
End Sub

'Buy Drums Online (OnlineDrums.vbp)
'frmCymbals (frmCymbals)
'Ben Harper
'3/23/06
'this form allows the user to shop cymbals through pictures and get general information about the product.
'This form also allows the user to select different sizes of the same cymbal through an input box and an array.
'In this form, the total for the cymbals alone will also be displayed.






Private Sub Z5_Click()    'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian Oriental China", FormatCurrency(149.49, 2)
    Cymbalsum = Cymbalsum + 149.49
End Sub

Private Sub Z6_Click()    'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian Custom K Crash", FormatCurrency(179.8, 2)
    Cymbalsum = Cymbalsum + 179.8
End Sub

Private Sub Z7_Click()
    Dim InchesArray(1 To 30) As Integer
    Dim NamesArray(1 To 12) As String
    Dim PricesArray(1 To 400) As Single
    Dim Found As Boolean
    Dim Pos, Search As Integer
        
    Open App.Path & "\ZildjianARide.txt" For Input As #1  'opens file with cymbal info
        Pos = 0
    Do Until EOF(1)
        Pos = Pos + 1
        Input #1, InchesArray(Pos), NamesArray(Pos), PricesArray(Pos) 'puts file into 3 arrays
    Loop
    Close #1
    
    Found = False
    Pos = 0
    Search = InputBox("enter the size (20-24 inches) of cymbal you would like", "Cymbal Size")
                                                                 'user inputs the size he would like
    Do While Found = False
        Pos = Pos + 1
        If Search = InchesArray(Pos) Then
        Found = True
        End If
    Loop
    If Found = True Then                                     'when the right size is found, put in cart and add to total cost
        picResultscymbals.Print NamesArray(Pos), FormatCurrency(PricesArray(Pos))
        Cymbalsum = Cymbalsum + PricesArray(Pos)
    End If
End Sub

Private Sub Z8_Click()       'adds cymbal to cymbal total
    picResultscymbals.Print "Zildjian Custom K Hi-Hat", FormatCurrency(212.99, 2)
    Cymbalsum = Cymbalsum + 212.99
End Sub

