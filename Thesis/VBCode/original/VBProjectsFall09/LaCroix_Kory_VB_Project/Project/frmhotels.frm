VERSION 5.00
Begin VB.Form frmhotels 
   BackColor       =   &H00400040&
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblHilton 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Text            =   "#2"
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdchoosehotel 
      Caption         =   "Choose A Hotel"
      Height          =   855
      Left            =   1440
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   5655
      Left            =   6240
      ScaleHeight     =   5595
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
   End
   Begin VB.PictureBox picSuperEight 
      Height          =   2295
      Left            =   3960
      Picture         =   "frmhotels.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picRamada 
      Height          =   855
      Left            =   3600
      Picture         =   "frmhotels.frx":1671
      ScaleHeight     =   795
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox PicHilton 
      Height          =   1455
      Left            =   720
      Picture         =   "frmhotels.frx":1F5F
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox piccomfortsuites 
      Height          =   1815
      Left            =   720
      Picture         =   "frmhotels.frx":2D7A
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdviewhotels 
      Caption         =   "View Hotels and Hotel Prices"
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "#4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblRamada 
      BackColor       =   &H00400040&
      Caption         =   "#3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblcomfort 
      BackColor       =   &H00400040&
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblhotels 
      BackColor       =   &H00400040&
      Caption         =   "Once You've arrived at your destination you will need to choose a hotel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmhotels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Brett Favre Fan Club
'Form Name: frmhotels
'Author: Kory LaCroix
'Date Written: 10/19/08
'Objective: To select a hotel
Private Sub cmdchoosehotel_Click()
Dim hotelchoice As Integer
'this asks the user for a number which is connect to a specific hotel.
'when the user enters a correct number the program moves to the next screen
'adds the cost of the hotel to the running total
'if the user selects an incorrect number it does not move on and ask for another number
hotelchoice = InputBox("Please enter the number of the hotel that you would like to stay in.")
    If hotelchoice = 1 Then
        MsgBox ("You have chosen Comfort Suites which will cost you $125.00.")
        runningtotal = runningtotal + 125
        frmhotels.Visible = False
        frmfinal.Visible = True
    ElseIf hotelchoice = 2 Then
        MsgBox ("You have chosen Hilton which will cost you $190.00.")
        runningtotal = runningtotal + 190
        frmhotels.Visible = False
        frmfinal.Visible = True
    ElseIf hotelchoice = 3 Then
        MsgBox ("You have chosen the Ramada which will cost you $300.00.")
        runningtotal = runningtotal + 300
        frmhotels.Visible = False
        frmfinal.Visible = True
    ElseIf hotelchoice = 4 Then
        MsgBox ("You have chosen the Super 8 which will cost you $105.00.")
        runningtotal = runningtotal + 105
        frmhotels.Visible = False
        frmfinal.Visible = True
    Else
        MsgBox ("You have entered an incorrect number. Please try again.")
    End If
End Sub

Private Sub cmdviewhotels_Click()
Dim hotel(1 To 50) As String
Dim hotelcost(1 To 50) As Double
'this following commands will make hte pictures of the hotels appaer
piccomfortsuites.Visible = True
PicHilton.Visible = True
cmdchoosehotel.Visible = True
picRamada.Visible = True
picSuperEight.Visible = True
cmdviewhotels.Visible = False

CTR = 0

'the following opens up a file containing the costs of the hotels
Open App.Path & "\Hotels.txt" For Input As #3

Do While Not EOF(3)
    CTR = CTR + 1
    Input #3, hotel(CTR), hotelcost(CTR)
Loop

picResults.Print "Hotel"; Tab(40); "Hotel Cost"
picResults.Print "******************************************************************"

For j = 1 To CTR
    picResults.Print hotel(j); Tab(40); FormatCurrency(hotelcost(j))
Next j

End Sub

