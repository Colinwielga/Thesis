VERSION 5.00
Begin VB.Form frmActivitiesPC 
   BackColor       =   &H00FFFF80&
   Caption         =   "Activities for ParkCentral"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3255
      Left            =   3600
      ScaleHeight     =   3195
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here to see a list of prices for these activities listed from the least to the greatest"
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBackPc 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click to go back to Rooms for Park Central"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   8160
      Picture         =   "frmActivities.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   240
      Picture         =   "frmActivities.frx":1012
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   7320
      Picture         =   "frmActivities.frx":202D
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   720
      Picture         =   "frmActivities.frx":C1E5
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblbaseball 
      BackColor       =   &H00FF8080&
      Caption         =   "Take in an afternoon baseball game"
      BeginProperty Font 
         Name            =   "GulimChe"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblspa 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enjoy a relaxing day at the spa!!"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblShow 
      BackColor       =   &H000080FF&
      Caption         =   "See a show on broadway!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lbltour 
      BackColor       =   &H008080FF&
      Caption         =   "Tour through Central Park!"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "frmActivitiesPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form allows the user to see some of the activities that they can
'do while staying at the Park Central hotel in New York City


Private Sub cmdBackPc_Click()
frmActivitiesPC.Hide
frmRoomPC.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdSort_Click()
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
