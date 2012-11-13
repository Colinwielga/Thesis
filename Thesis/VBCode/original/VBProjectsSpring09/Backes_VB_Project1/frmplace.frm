VERSION 5.00
Begin VB.Form frmactivitiesWarwick 
   BackColor       =   &H00C0FFFF&
   Caption         =   "frm activities for the Warwick"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   FillColor       =   &H00400000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3495
      Left            =   3480
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FF80FF&
      Caption         =   "Click here to see a list of prices for the activities listed from the least to the greatest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackWarwick 
      BackColor       =   &H0080C0FF&
      Caption         =   "Click to go back to the Rooms for the Warwick"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   8040
      Picture         =   "frmplace.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   360
      Picture         =   "frmplace.frx":0BF6
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   7560
      Picture         =   "frmplace.frx":1942
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   600
      Picture         =   "frmplace.frx":2BB3
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblbaseball 
      BackColor       =   &H00800000&
      Caption         =   "Check out a baseball game at Yankee Stadium"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7200
      TabIndex        =   6
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblchill 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Take the day off...and go to the SPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lblshow 
      BackColor       =   &H00C000C0&
      Caption         =   "Take in a show on broadway!!"
      Height          =   1335
      Left            =   9000
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblcentral 
      BackColor       =   &H00FFFF00&
      Caption         =   "Take a tour through the beautiful Central Park"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmactivitiesWarwick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user about the different activities
'they can participate in while staying at the Warwick
Option Explicit

Private Sub cmdBackWarwick_Click()
'allows the user to go back to the room selection form
frmactivitiesWarwick.Hide
frmRoomsWarwick.Show
End Sub

Private Sub cmdquit_Click()
'Quits the form
End
End Sub

Private Sub cmdSort_Click()
'declare all your variables
Dim pass As Integer, pos As Integer, tempActivities As String
Dim tempPriceList As Single, Activities(1 To 4) As String
Dim PriceList(1 To 4) As Single
CTR = 0
'open the file where information is
Open App.Path & "\NYactivityPrices.txt" For Input As #1
'put the file into an array
Do While Not EOF(1)
   CTR = CTR + 1
   Input #1, Activities(CTR), PriceList(CTR)
Loop
Close #1

'sorts the prices from least to greatest
For pass = 1 To CTR - 1
    For pos = 1 To CTR - pass
        If PriceList(pos) > PriceList(pos + 1) Then
           tempPriceList = PriceList(pos)
           PriceList(pos) = PriceList(pos + 1)
           PriceList(pos + 1) = tempPriceList
           
           tempActivities = Activities(pos)
           Activities(pos) = Activities(pos + 1)
           Activities(pos + 1) = tempActivities
           
End If
  Next pos
    Next pass
    
picResults.Print "Activity", "Price"
picResults.Print "*******************************************"

For X = 1 To CTR
    picResults.Print Activities(X), FormatCurrency(PriceList(X))
Next X
End Sub
