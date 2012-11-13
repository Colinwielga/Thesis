VERSION 5.00
Begin VB.Form frmactivitiesHilton 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Things to do while at the Hilton"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   3855
      Left            =   3480
      ScaleHeight     =   3795
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H008080FF&
      Caption         =   "Click to see a list of prices for these activities from ranging from the least to the greatest"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FF80&
      Caption         =   "Click to go back to the rooms for the Hilton"
      BeginProperty Font 
         Name            =   "Mathematica5"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   8040
      Picture         =   "frmDoSomething.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   5160
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   960
      Picture         =   "frmDoSomething.frx":1020
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   7800
      Picture         =   "frmDoSomething.frx":219A
      ScaleHeight     =   1875
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   360
      Picture         =   "frmDoSomething.frx":2F92
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblshop 
      BackColor       =   &H00800080&
      Caption         =   "Like shopping? Check out some of our stores!"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lbltour 
      BackColor       =   &H00FFFF00&
      Caption         =   "Tour our wonderful city!"
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblgolf 
      BackColor       =   &H0000FF00&
      Caption         =   "Hit the links on one of our many golf courses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   9000
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblrelax 
      BackColor       =   &H0000FFFF&
      Caption         =   "Relax you are on vacation....why not go to the SPA?!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmactivitiesHilton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objectives: This form informs the user of activities that they
'could do during their stay at the Hilton

Private Sub cmdBack_Click()
'allows the user to go back to the room selection form
frmRoomsHilton.Show
frmactivitiesHilton.Hide

End Sub

Private Sub cmdquit_Click()
'closes the form
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
