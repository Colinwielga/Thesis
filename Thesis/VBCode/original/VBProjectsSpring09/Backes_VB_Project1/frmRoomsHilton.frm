VERSION 5.00
Begin VB.Form frmRoomsHilton 
   BackColor       =   &H00FF00FF&
   Caption         =   "Rooms for the Hilton"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdactivities 
      BackColor       =   &H00FF80FF&
      Caption         =   "Click to see some Activities in LA"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdRestaurants 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to see some Restaurant options at the Hilton"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to go back to the previous page"
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FF8080&
      Caption         =   "Click here to view the room options"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   4695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H0000FF00&
      Caption         =   "Click to see your Total"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      Height          =   4575
      Left            =   2640
      ScaleHeight     =   4515
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
   Begin VB.TextBox txtRnumber 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtnights 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtguests 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblRnumber 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number for the room you would like to stay in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblnights 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter the number of nights you would like to stay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblguests 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter Number of guests that will be staying"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmRoomsHilton"
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

Private Sub cmdactivities_Click()
'allows the user to go to the activities form
frmRoomsHilton.Hide
frmactivitiesHilton.Show

End Sub

Private Sub cmdBack_Click()
'allows the user to go to the home form for the hilton
frmRoomsHilton.Hide
frmHilton.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub


Private Sub cmdRestaurants_Click()
'allows the user to go to the restaurants form
frmRoomsHilton.Hide
frmRestaurantsHilton.Show

End Sub

Private Sub cmdView_Click()
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
