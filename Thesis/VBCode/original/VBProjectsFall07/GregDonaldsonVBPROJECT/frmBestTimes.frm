VERSION 5.00
Begin VB.Form frmBestTimes 
   BackColor       =   &H00FF0000&
   Caption         =   "Best Times"
   ClientHeight    =   5235
   ClientLeft      =   3840
   ClientTop       =   3060
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8070
   Begin VB.PictureBox picTimes 
      Height          =   4095
      Left            =   2880
      ScaleHeight     =   4035
      ScaleWidth      =   4155
      TabIndex        =   15
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdOneBreast 
      Caption         =   "100 Breaststroke"
      Height          =   615
      Left            =   1440
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOneFree 
      Caption         =   "100 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdThousand 
      Caption         =   "1000 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "500 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdTwoFree 
      Caption         =   "200 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdFifty 
      Caption         =   "50 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOneBack 
      Caption         =   "100 Backstroke"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTwoBack 
      Caption         =   "200 Backstroke"
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOneFly 
      Caption         =   "100 Butterfly"
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdTwoFly 
      Caption         =   "200 Butterfly"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdTwoBreast 
      Caption         =   "200 Breaststroke"
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdTwoIM 
      Caption         =   "200 IM"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdFourIM 
      Caption         =   "400 IM"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to the Home Page"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdMile 
      Caption         =   "1650 Freestyle"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBestTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFifty_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\50Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdFive_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\500Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdFourIM_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\400IM.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdMile_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
Dim temp As String, comp As Integer

picTimes.Cls
Ctr = 0
Open App.Path & "\1650.txt" For Input As #1

Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdOneBack_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\100Back.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdOneBreast_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\100Breast.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
        picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdOneFly_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\100Fly.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdOneFree_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\100Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdReturn_Click()
'hides the best times form and returns to the home page
frmBestTimes.Hide
frmHomePage.Show
End Sub

Private Sub cmdThousand_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
Dim temp As String, comp As Integer
Dim Pass, Pos As Integer

picTimes.Cls
Ctr = 0
Open App.Path & "\1000.txt" For Input As #1

Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdTwoBack_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\200Back.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdTwoBreast_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\200Breast.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdTwoFly_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\200Fly.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdTwoFree_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\200Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub

Private Sub cmdTwoIM_Click()
'Opens file into an array and displays contents

Dim Swimmer(1 To 100) As String
Dim Time(1 To 100) As String
picTimes.Cls
Ctr = 0
Open App.Path & "\200IM.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For J = 1 To Ctr
    picTimes.Print Swimmer(J); Tab(15); Time(J)
Next J
Close #1
End Sub
