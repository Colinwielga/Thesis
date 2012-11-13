VERSION 5.00
Begin VB.Form frmActivitiesMarriott 
   BackColor       =   &H0080FFFF&
   Caption         =   "things to do while staying at the Marriott"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3015
      Left            =   3720
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   11
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click to see a list of prices for these activities, sorted from the least to the greatest"
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to go back to Rooms"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   7080
      Picture         =   "frmGetOutThere.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   7560
      Picture         =   "frmGetOutThere.frx":0E10
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Left            =   120
      Picture         =   "frmGetOutThere.frx":1B64
      ScaleHeight     =   1635
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   1200
      Picture         =   "frmGetOutThere.frx":295D
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblshop 
      BackColor       =   &H00C00000&
      Caption         =   "Go Shopping at some of our fabulous stores "
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   8760
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblgolf 
      BackColor       =   &H00C000C0&
      Caption         =   "Go play a round of golf at one of our many beautiful courses!"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lbltour 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Take a tour of our wonderful city!"
      BeginProperty Font 
         Name            =   "MingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblspa 
      BackColor       =   &H00FFFF00&
      Caption         =   "Treat yourself to a day at the SPA!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmActivitiesMarriott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user about the different activities
'they can participate in while staying at the Marriott

Private Sub cmdBack_Click()
'allows the user to go back to the room selection form
frmActivitiesMarriott.Hide
frmRoomMarriott.Show

End Sub

Private Sub cmdquit_Click()
'Quits the form
End
End Sub

Private Sub cmdSort_Click()
Dim pass As Integer, pos As Integer, tempActivities As String
Dim tempPriceList As Single, Activities(1 To 4) As String
Dim PriceList(1 To 4) As Single
CTR = 0
'open the file where information is
Open App.Path & "\LAactivityPrices.txt" For Input As #1
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
