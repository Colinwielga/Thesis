VERSION 5.00
Begin VB.Form frmNewTime 
   BackColor       =   &H00FF0000&
   Caption         =   "New Time"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results Box"
      Height          =   615
      Left            =   3120
      TabIndex        =   19
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return to Home Page"
      Height          =   615
      Left            =   6600
      TabIndex        =   18
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   3480
      ScaleHeight     =   3555
      ScaleWidth      =   3435
      TabIndex        =   16
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtNewTime 
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOneBreast 
      Caption         =   "100 Breaststroke"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOneFly 
      Caption         =   "100 Butterfly"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwoFree 
      Caption         =   "200 Freestyle"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwoIM 
      Caption         =   "200 IM"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdThou 
      Caption         =   "1000 Freestyle"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOneFree 
      Caption         =   "100 Freestyle"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwoBack 
      Caption         =   "200 Backstroke"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwoBreast 
      Caption         =   "200 Breaststroke"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdTwoFly 
      Caption         =   "200 Butterfly"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdFiveFree 
      Caption         =   "500 Freestyle"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdFourIM 
      Caption         =   "400 IM"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdMile 
      Caption         =   "1650 Freestyle"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOneBack 
      Caption         =   "100 Backstroke"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdFifty 
      Caption         =   "50 Freestyle"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter time like example{ 2:29.56 or for times under one minute{ 0:46.38"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5760
      TabIndex        =   17
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "ENTER TIME FIRST IN BOX ON RIGHT, THEN SELECT THE EVENT TO COMPARE IT TO."
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmNewTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
picResults.Cls
End Sub

'the following sub routines compare a new time entered by the user to existing
'times already stored in the computer. this program only compares the fastest time
'to the new time to determine if a new fastest time has been entered.
Private Sub cmdFifty_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text  'assigns textbox value to NewTime
Open App.Path & "\50Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1    'sorts open file to compare fastest time with the new time
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdFiveFree_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\500Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdFourIM_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\400IM.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdMile_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\1650.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdOneBack_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\100Back.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdOneBreast_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\100Breast.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdOneFly_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\100Fly.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdOneFree_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\100Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdReturnHome_Click()
    frmNewTime.Hide
    frmHomePage.Show
End Sub

Private Sub cmdThou_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\1000.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdTwoBack_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\200Back.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdTwoBreast_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\200Breast.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdTwoFly_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\200Fly.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdTwoFree_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\200Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub

Private Sub cmdTwoIM_Click()
Dim Ctr As Integer
Dim NewTime As String
Dim Swimmer(1 To 100), Time(1 To 100) As String

NewTime = txtNewTime.Text
Open App.Path & "\200IM.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), Time(Ctr)
Loop

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Time(Pos) > Time(Pos + 1) Then
            temp = Time(Pos)
            Time(Pos) = Time(Pos + 1)
            Time(Pos + 1) = temp
        End If
    Next Pos
Next Pass

If NewTime < Time(1) Then
    picResults.Print "New fastest time!"
Else
    picResults.Print "Sorry, the new time is not the fastest."
End If
Close #1
End Sub
