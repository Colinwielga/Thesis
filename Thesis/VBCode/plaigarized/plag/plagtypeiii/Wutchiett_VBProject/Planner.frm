VERSION 5.00
Begin VB.Form Weather
   Caption         =   "Form8"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form8"
   Picture         =   "Planner.frx":0000
   ScaleHeight     =   5535
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5
      Caption         =   "Minnesota Centennial Showboat"
      Height          =   615
      Left            =   6840
      TabIndex        =   14
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command4
      Caption         =   "James J Hill Museum"
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command3
      Caption         =   "Como Park Zoo and Conservatory"
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox Picture1
      Height          =   3495
      Left            =   3840
      ScaleHeight     =   3435
      ScaleWidth      =   5115
      TabIndex        =   11
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton Command2
      Caption         =   "Weather Forcast"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox picResults
      Height          =   3255
      Left            =   3840
      Picture         =   "Planner.frx":AFCC2
      ScaleHeight     =   3195
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.CommandButton Command1
      Caption         =   "Return"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label8
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Attractions"
      BeginProperty Font
         Name            =   "Bradley Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2160
      TabIndex        =   15
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label Label7
      BackStyle       =   0  'Transparent
      Caption         =   "7. Saturday"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label6
      BackStyle       =   0  'Transparent
      Caption         =   "6. Friday"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5
      BackStyle       =   0  'Transparent
      Caption         =   "5. Thursday"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4
      BackStyle       =   0  'Transparent
      Caption         =   "4. Wednesday"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3
      BackStyle       =   0  'Transparent
      Caption         =   "3. Tuesday"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2
      BackStyle       =   0  'Transparent
      Caption         =   "2. Monday"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "1. Sunday"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label a
      BackStyle       =   0  'Transparent
      Caption         =   "Weather"
      BeginProperty Font
         Name            =   "Bradley Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Weather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St.PaulEvents
'Weather
'David Wutchiett
'February 24, 2010
'Checks weather through inputboxes, and provides descriptions of other attractions.


Option Explicit

Private Sub Command1_Click()
Form1.Show
Weather.Hide


End Sub

Private Sub Command2_Click()

Dim Day As Integer

Day = InputBox("Enter a number for the day of the week in which you are interested. (1-7)")

If Day = 7 Then
        MsgBox ("It will be slighly cloud with a chance of rain. Low of 32F, High of 45F.")
        ElseIf Day = 6 Then
        MsgBox ("70% chance of rain. Low of 38F, High of 43F.")
        ElseIf Day = 5 Then
        MsgBox ("Likely chances of snow. Low of 15F, High of 17F.")
        ElseIf Day = 4 Then
        MsgBox ("There will be a hurricane. Low of 77F, High of 88F.")
        ElseIf Day = 3 Then
        MsgBox ("Desert conditions, wear some sandals. Low of 102F, High of 145F.")
        ElseIf Day = 2 Then
        MsgBox ("Blizzard. Low of 22F, High of 27F.")
        ElseIf Day = 1 Then
        MsgBox ("Sunny! Low of 39F, High of 48F.")
        Else
        MsgBox ("Please enter a number 1-7")
    End If



End Sub

Private Sub Command3_Click()

Picture1.ForeColor = vbBlue

Picture1.Print "Como Park Zoo and Conservatory"
Picture1.Print "Located in historic Como Park, this popular zoo, especially known for its"
Picture1.Print "California sea lion exhibit, also features a great cat display, gorillas"
Picture1.Print "and giraffes."
Picture1.Print " "
End Sub

Private Sub Weather_Click()

End Sub

Private Sub Command4_Click()

Picture1.ForeColor = vbBlue
Picture1.Print "James J Hill Museum"
Picture1.Print "This massive red sandstone mansion was the home of James J. Hill,"
Picture1.Print "builder of the Great Northern Railway, and is a fine representation of life "
Picture1.Print "during the Gilded Age."
Picture1.Print " "

End Sub

Private Sub Command5_Click()

Picture1.ForeColor = vbBlue
Picture1.Print "Minnesota Centennial Showboat"
Picture1.Print "The Minnesota Centennial Showboat, docked on the banks of the"
Picture1.Print "Mighty Mississippi, makes old time theatre big time fun with its turn-of-"
Picture1.Print "the-century interior, vaudevillian entertainment, and timeless tales."
Picture1.Print "Enjoy dinner or a river cruise with your show to enhance this"
Picture1.Print "extraordinary theatre experience. Come relive the whimsical times of"
Picture1.Print "long ago on the Showboat."
Picture1.Print " "
End Sub

