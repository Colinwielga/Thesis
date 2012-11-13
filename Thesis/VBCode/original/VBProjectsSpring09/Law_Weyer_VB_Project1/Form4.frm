VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form4"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form4"
   ScaleHeight     =   6315
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpend 
      Caption         =   "Book Flight"
      Height          =   615
      Left            =   6480
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtSpend 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input Flight List"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   4455
      Left            =   2280
      ScaleHeight     =   4395
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Destination"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSortPrice 
      Caption         =   "Sort By Price"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Travel Times"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblSpend 
      BackStyle       =   0  'Transparent
      Caption         =   "How much would you like to spend, in exact dollars?"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ctr As Integer
Dim Place(1 To 100) As String
Dim Price(1 To 100) As Single
Dim DepartTime(1 To 100) As String
Dim ArrivalTime(1 To 100) As String
Dim i As Integer
Dim Pass As Integer, Pos As Integer
Dim Temp As Integer

Private Sub cmdCancel_Click()
Form4.Hide
Form1.Show
End Sub

Private Sub cmdInput_Click()

Open App.Path & "\flight.txt" For Input As #1

Ctr = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Place(Ctr), Price(Ctr), DepartTime(Ctr), ArrivalTime(Ctr)
Loop
Close #1

End Sub

Private Sub cmdPrint_Click()


picResults.Cls
picResults.Print "Destination"; Tab(25); "Ticket Price", Tab(45); " Departure Time"; Tab(65); "Arrival Time"
picResults.Print "**************************************************************************************************"
For i = 1 To Ctr
picResults.Print Place(i); Tab(25); FormatCurrency(Price(i)); Tab(45); DepartTime(i); Tab(65); ArrivalTime(i)

Next i

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSearch_Click()
Dim Found As Boolean
Dim PlaceCtr As Integer, WantPlace As String
Found = False
Dim K As Integer

Open App.Path & "\flight.txt" For Input As #1

PlaceCtr = 0
Do While Not EOF(1)
    PlaceCtr = PlaceCtr + 1
    Input #1, Place(PlaceCtr), Price(PlaceCtr), DepartTime(PlaceCtr), ArrivalTime(PlaceCtr)
Loop
Close #1

picResults.Cls

WantPlace = InputBox("Enter the destination that you want to go to.", "Destination")

picResults.Print "Destination"; Tab(25); "Ticket Price", Tab(45); " Departure Time"; Tab(65); "Arrival Time"
picResults.Print "**************************************************************************************************"

For K = 1 To PlaceCtr
    If WantPlace = Place(K) Then
      Found = True
      Ctr = Ctr + 1
    End If
    
    If WantPlace <> Place(K) Then
      Found = False
    End If

If Found Then
    picResults.Print Place(K); Tab(25); FormatCurrency(Price(K)); Tab(45); DepartTime(K); Tab(65); ArrivalTime(K)
End If
Next K

If Ctr < 1 Then
    picResults.Print WantPlace; " is not in the database."
End If
Ctr = 0

End Sub

Private Sub cmdSortPrice_Click()
Dim TempPlace As String, TempPrice As Single, TempDeparttime As String, TempArrivaltime As String

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Price(Pos) > Price(Pos + 1) Then
            TempPlace = Place(Pos)
            Place(Pos) = Place(Pos + 1)
            Place(Pos + 1) = TempPlace
            
            TempPrice = Price(Pos)
            Price(Pos) = Price(Pos + 1)
            Price(Pos + 1) = TempPrice
            
            TempDeparttime = DepartTime(Pos)
            DepartTime(Pos) = DepartTime(Pos + 1)
            DepartTime(Pos + 1) = TempDeparttime
            
            TempArrivaltime = ArrivalTime(Pos)
            ArrivalTime(Pos) = ArrivalTime(Pos + 1)
            ArrivalTime(Pos + 1) = TempArrivaltime
        End If
    Next Pos
Next Pass

picResults.Cls
picResults.Print "Destination"; Tab(25); "Ticket Price", Tab(45); " Departure Time"; Tab(65); "Arrival Time"
picResults.Print "**************************************************************************************************"
For i = 1 To Ctr
picResults.Print Place(i); Tab(25); FormatCurrency(Price(i)); Tab(45); DepartTime(i); Tab(65); ArrivalTime(i)

Next i
End Sub

Private Sub cmdSpend_Click()
Dim Found As Boolean
Dim PlaceCtr As Integer
Found = False
Dim K As Integer, Spend As Integer

Open App.Path & "\flight.txt" For Input As #1

PlaceCtr = 0
Do While Not EOF(1)
    PlaceCtr = PlaceCtr + 1
    Input #1, Place(PlaceCtr), Price(PlaceCtr), DepartTime(PlaceCtr), ArrivalTime(PlaceCtr)
Loop
Close #1

Spend = txtSpend.Text

For K = 1 To PlaceCtr
    If Spend = Price(K) Then                'Does a search and stop which finds the flight the user wants to book
      Found = True
    End If
    
    If Spend <> Price(K) Then
      Found = False
    End If

    If Found Then
        Destination = Place(K)
        Budget = Price(K)
        DepartureTicketTime = DepartTime(K)
        ArrivalTicketTime = ArrivalTime(K)
    End If
Next K

Form4.Hide                      'Switches back to original form as this one is done with
Form1.Show

End Sub
