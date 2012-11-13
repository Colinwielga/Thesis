VERSION 5.00
Begin VB.Form frmRoomMarriott 
   BackColor       =   &H00FFFF00&
   Caption         =   "Marriott Rooms"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestaurants 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click here to see your Restaurant options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "click here to check out some activities"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF8080&
      Caption         =   "Click to go back to previous page"
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdsize 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here to load the room options"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FF80FF&
      Caption         =   "Total for Room"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@MS PGothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtRnumber 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox txtnights 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtguests 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   2760
      ScaleHeight     =   3915
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lbltype 
      BackColor       =   &H00C000C0&
      Caption         =   "Enter the number of which room you would like to stay in"
      BeginProperty Font 
         Name            =   "@Gulim"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblnights 
      BackColor       =   &H00C00000&
      Caption         =   "Enter Number of night that you will be staying"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblguests 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter number of guests that will be coming"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmRoomMarriott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form allows the user to enter in how many guests
'will be staying, at also allows them to see the different room
'options and tells them their total for their room

Option Explicit

Private Sub cmdBack_Click()
'allows the user to go back to the home form for the Marriott
frmRoomMarriott.Hide
frmMarriott.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub


Private Sub cmdRestaurants_Click()
'allows the user to go to the restaurants page
frmRoomMarriott.Hide
frmRestaurantsMarriott.Show


End Sub

Private Sub cmdsize_Click()
Dim NUMBER(1 To 4) As Integer, ROOMTYPE(1 To 4) As String
Dim COST(1 To 4) As Single, BEDS(1 To 4) As String
'get the file ready
Open App.Path & "\Rooms.txt" For Input As #1

'put the file into an array
Do While Not EOF(1)
   CTR = CTR + 1
   Input #1, BEDS(CTR)
   Input #1, NUMBER(CTR)
   Input #1, ROOMTYPE(CTR)
   Input #1, COST(CTR)
Loop
Close #1
'table for the picture box
picResults.Print "Number", "Room Selection", "Cost"
picResults.Print "***********************************************************"

For X = 1 To CTR
   picResults.Print BEDS(X), ROOMTYPE(X), FormatCurrency(COST(X))
Next X
cmdTotal.Enabled = True

End Sub

Private Sub cmdTotal_Click()
Dim guests As Integer, nights As Integer, Rnumber As Integer
Dim total As Single, tax As Single


guests = txtguests.Text
nights = txtnights.Text
Rnumber = txtRnumber.Text

'figure out the total and tax for the different rooms chosen
If Rnumber = 1 Then
   total = nights * 135
ElseIf Rnumber = 2 Then
   total = nights * 140
ElseIf Rnumber = 3 Then
   total = nights * 150
ElseIf Rnumber = 4 Then
   total = nights * 160
End If

'now i added the tax to the total
tax = 0.3 * total

picResults.Print "Your total for your room is"; FormatCurrency(total)
End Sub

Private Sub Command1_Click()
'allows the user to go to the activities page
frmRoomMarriott.Hide
frmActivitiesMarriott.Show
End Sub


