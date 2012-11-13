VERSION 5.00
Begin VB.Form frmIreland1 
   BackColor       =   &H00008000&
   Caption         =   "Ireland 1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H000080FF&
      Height          =   4815
      Left            =   4920
      ScaleHeight     =   4755
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   2160
      Width           =   5415
   End
   Begin VB.CommandButton cmdHotelName 
      Caption         =   "Sort By Hotel Name"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Your Housing"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdMoreStars 
      Caption         =   "Sort by Stars"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Check out some hotel pricing for the Top Cities! (In Dollars)"
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmsPrices 
      Caption         =   "Take a look at Some Scereny Here!"
      Height          =   975
      Left            =   5520
      TabIndex        =   5
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuiz 
      Caption         =   "Take a Fun Fact Quiz Here! "
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdTravelABC 
      Caption         =   "Sort By ABC"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdTravel 
      BackColor       =   &H00C00000&
      Caption         =   "Click Here to see some of the top places to travel in Ireland"
      Height          =   1095
      Left            =   360
      MaskColor       =   &H0000FF00&
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label I 
      BackColor       =   &H00008000&
      Caption         =   "Information On the Beautiful Country of Ireland!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmIreland1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Information About Ireland
'Form Name: Ireland1
'Author: Rachel Lietzke
'Date Written: March 27, 2008
'Objective: To Offer Best Places to Travel and Where to Stay
'Also this is the Main page and will allow you to access any other form from it

Private Sub cmdCalculate_Click()
frmIreland1.Hide
frmIreland4.Show
End Sub

Private Sub cmdHotelName_Click(Index As Integer)
Dim City(1 To 30) As String
Dim Names(1 To 30) As String
Dim Stars(1 To 30) As Integer
Dim Cost(1 To 30) As Single
Dim CTR As Integer
Dim POS As Integer
Dim PASS As Integer
Dim J As Integer
Dim TempName As String

CTR = 0

Open App.Path & "\More.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, City(CTR), Names(CTR), Stars(CTR), Cost(CTR)
Loop
picResults.Cls
picResults.Print "Hotel"; Tab(30); "City"; Tab(40); "Stars"; Tab(50); "Price"
picResults.Print "**************************************************************************"

For PASS = 1 To CTR
    For POS = 1 To CTR - PASS
        If Names(POS) > Names(POS + 1) Then
            TempName = Names(POS)
            Names(POS) = Names(POS + 1)
            Names(POS + 1) = TempName
            TempName = City(POS)
            City(POS) = City(POS + 1)
            City(POS + 1) = TempName
            TempName = Stars(POS)
            Stars(POS) = Stars(POS + 1)
            Stars(POS + 1) = TempName
            TempName = Cost(POS)
            Cost(POS) = Cost(POS + 1)
            Cost(POS + 1) = TempName
        End If
        
    Next POS
Next PASS

For J = 1 To CTR
    picResults.Print Names(J); Tab(30); City(J); Tab(40); Stars(J); Tab(50); Cost(J)
Next J
Close #1

End Sub

Private Sub cmdMore_Click()
Dim City(1 To 30) As String
Dim Name(1 To 30) As String
Dim Stars(1 To 30) As Integer
Dim Cost(1 To 30) As Single
Dim CTR As Integer
CTR = 0
Open App.Path & "\More.txt" For Input As #1
picResults.Cls
picResults.Print "City"; Tab(10); "Hotel"; Tab(40); "Stars"; Tab(50); "Price"
picResults.Print "**************************************************************************"

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, City(CTR), Name(CTR), Stars(CTR), Cost(CTR)
    picResults.Print City(CTR); Tab(10); Name(CTR); Tab(40); Stars(CTR); Tab(50); Cost(CTR)
Loop
Close #1
End Sub

Private Sub cmdMoreStars_Click(Index As Integer)
Dim City(1 To 30) As String
Dim Names(1 To 30) As String
Dim Stars(1 To 30) As Integer
Dim Cost(1 To 30) As Single
Dim CTR As Integer
Dim POS As Integer
Dim PASS As Integer
Dim J As Integer
Dim TempName As String

CTR = 0

Open App.Path & "\More.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, City(CTR), Names(CTR), Stars(CTR), Cost(CTR)
Loop
picResults.Cls
picResults.Print "Stars"; Tab(10); "City"; Tab(20); "Hotel"; Tab(50); "Price"
picResults.Print "**************************************************************************"

For PASS = 1 To CTR
    For POS = 1 To CTR - PASS
        If Stars(POS) > Stars(POS + 1) Then
            TempName = Stars(POS)
            Stars(POS) = Stars(POS + 1)
            Stars(POS + 1) = TempName
            TempName = City(POS)
            City(POS) = City(POS + 1)
            City(POS + 1) = TempName
            TempName = Names(POS)
            Names(POS) = Names(POS + 1)
            Names(POS + 1) = TempName
            TempName = Cost(POS)
            Cost(POS) = Cost(POS + 1)
            Cost(POS + 1) = TempName
        End If
        
    Next POS
Next PASS

For J = 1 To CTR
    picResults.Print Stars(J); Tab(10); City(J); Tab(20); Names(J); Tab(50); Cost(J)
Next J
Close #1
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdQuiz_Click()
frmIreland1.Hide
frmIreland2.Show
End Sub

Private Sub cmdTravel_Click()
Dim Places(1 To 15) As String
Dim Position(1 To 15) As Integer
Dim CTR As Integer
CTR = 0
Open App.Path & "\TopPlaces.txt" For Input As #1

picResults.Cls
picResults.Print "Top Places to Visit"
picResults.Print "**********************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Position(CTR), Places(CTR)
    picResults.Print Position(CTR); Tab(20); Places(CTR)
Loop
Close #1
End Sub

Private Sub cmdTravelABC_Click(Index As Integer)
Dim Places(1 To 15) As String
Dim Position(1 To 15) As Integer
Dim CTR As Integer
Dim POS As Integer
Dim PASS As Integer
Dim J As Integer
Dim TempName As String
CTR = 0
Open App.Path & "\TopPlaces.txt" For Input As #1

picResults.Cls
picResults.Print "Top Places to Visit"
picResults.Print "**********************************"

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Position(CTR), Places(CTR)
Loop

For PASS = 1 To CTR
    For POS = 1 To CTR - PASS
        If Places(POS) > Places(POS + 1) Then
            TempName = Places(POS)
            Places(POS) = Places(POS + 1)
            Places(POS + 1) = TempName
            TempName = Position(POS)
            Position(POS) = Position(POS + 1)
            Position(POS + 1) = TempName
        End If
        
    Next POS
Next PASS

For J = 1 To CTR
    picResults.Print Places(J); Tab(20); Position(J)
Next J
Close #1

End Sub

Private Sub cmsPrices_Click()
frmIreland1.Hide
frmIreland3.Show
End Sub

