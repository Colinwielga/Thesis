VERSION 5.00
Begin VB.Form frmIndividualSearch 
   BackColor       =   &H000000FF&
   Caption         =   "Individual Search"
   ClientHeight    =   5130
   ClientLeft      =   3870
   ClientTop       =   3285
   ClientWidth     =   8025
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   8025
   Begin VB.CommandButton cmdReturnHome 
      Caption         =   "Return to Home Page"
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results Box"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   600
      ScaleHeight     =   2175
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   1920
      Width           =   6615
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF0000&
      Caption         =   "Search"
      Height          =   615
      Left            =   5520
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtIndividual 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Swimmers: Bobby, Clarence, Colin, Darren, Greg, Jeff, Josh, Justin, Karl, Kevin, Michael, Neil, Nick, Pat, Scott, Torri, Trent"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Enter name of swimmer to search for: (case sensitive, capitalize first letter only)"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmIndividualSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
picResults.Cls

End Sub

Private Sub cmdReturnHome_Click()
    frmIndividualSearch.Hide
    frmHomePage.Show
End Sub

Private Sub cmdSearch_Click()
'found this code setup in the text book on page 7[29].
'the following search multiple arrays for all the events that a
'requested swimmer swam
Dim Found As Boolean
Dim A, B As Integer
Dim SwimmerInput As String, Swimmer(1 To 100) As String
Dim SwimmerTime(1 To 100) As String

SwimmerInput = txtIndividual.Text
A = 0
Ctr = 0
Found = False
Open App.Path & "\50Free.txt" For Input As #1
Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (A < Ctr))
    A = A + 1
    If SwimmerInput = Swimmer(A) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 50 Freestyle with a time of "; SwimmerTime(A)
End If
Close #1

B = 0
Ctr = 0
Found = False
Open App.Path & "\100Back.txt" For Input As #2
Do Until EOF(2)
    Ctr = Ctr + 1
    Input #2, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (B < Ctr))
    B = B + 1
    If SwimmerInput = Swimmer(B) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 100 Backstroke with a time of "; SwimmerTime(B)
End If
Close #2

C = 0
Ctr = 0
Found = False
Open App.Path & "\100Breast.txt" For Input As #3
Do Until EOF(3)
    Ctr = Ctr + 1
    Input #3, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (C < Ctr))
    C = C + 1
    If SwimmerInput = Swimmer(C) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 100 Breaststroke with a time of "; SwimmerTime(C)
End If
Close #3

D = 0
Ctr = 0
Found = False
Open App.Path & "\100Fly.txt" For Input As #4
Do Until EOF(4)
    Ctr = Ctr + 1
    Input #4, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (D < Ctr))
    D = D + 1
    If SwimmerInput = Swimmer(D) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 100 Butterfly with a time of "; SwimmerTime(D)
End If
Close #4

e = 0
Ctr = 0
Found = False
Open App.Path & "\100Free.txt" For Input As #5
Do Until EOF(5)
    Ctr = Ctr + 1
    Input #5, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (e < Ctr))
    e = e + 1
    If SwimmerInput = Swimmer(e) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 100 Freestyle with a time of "; SwimmerTime(e)
End If
Close #5

F = 0
Ctr = 0
Found = False
Open App.Path & "\200Back.txt" For Input As #6
Do Until EOF(6)
    Ctr = Ctr + 1
    Input #6, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (F < Ctr))
    F = F + 1
    If SwimmerInput = Swimmer(F) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 200 Backstroke with a time of "; SwimmerTime(F)
End If
Close #6


G = 0
Ctr = 0
Found = False
Open App.Path & "\200Breast.txt" For Input As #7
Do Until EOF(7)
    Ctr = Ctr + 1
    Input #7, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (G < Ctr))
    G = G + 1
    If SwimmerInput = Swimmer(G) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 200 Breaststroke with a time of "; SwimmerTime(G)
End If
Close #7

H = 0
Ctr = 0
Found = False
Open App.Path & "\200Fly.txt" For Input As #8
Do Until EOF(8)
    Ctr = Ctr + 1
    Input #8, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (H < Ctr))
    H = H + 1
    If SwimmerInput = Swimmer(H) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 200 Butterfly with a time of "; SwimmerTime(H)
End If
Close #8

I = 0
Ctr = 0
Found = False
Open App.Path & "\200Free.txt" For Input As #9
Do Until EOF(9)
    Ctr = Ctr + 1
    Input #9, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (I < Ctr))
    I = I + 1
    If SwimmerInput = Swimmer(I) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 200 Freestyle with a time of "; SwimmerTime(I)
End If
Close #9

J = 0
Ctr = 0
Found = False
Open App.Path & "\200IM.txt" For Input As #10
Do Until EOF(10)
    Ctr = Ctr + 1
    Input #10, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (J < Ctr))
    J = J + 1
    If SwimmerInput = Swimmer(J) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 200 Individual Medley with a time of "; SwimmerTime(J)
End If
Close #10

K = 0
Ctr = 0
Found = False
Open App.Path & "\400IM.txt" For Input As #11
Do Until EOF(11)
    Ctr = Ctr + 1
    Input #11, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (K < Ctr))
    K = K + 1
    If SwimmerInput = Swimmer(K) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 400 Individual Medley with a time of "; SwimmerTime(K)
End If
Close #11

L = 0
Ctr = 0
Found = False
Open App.Path & "\500Free.txt" For Input As #12
Do Until EOF(12)
    Ctr = Ctr + 1
    Input #12, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (L < Ctr))
    L = L + 1
    If SwimmerInput = Swimmer(L) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 500 Freestyle with a time of "; SwimmerTime(L)
End If
Close #12

M = 0
Ctr = 0
Found = False
Open App.Path & "\1000.txt" For Input As #13
Do Until EOF(13)
    Ctr = Ctr + 1
    Input #13, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (M < Ctr))
    M = M + 1
    If SwimmerInput = Swimmer(M) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 1000 Freestyle with a time of "; SwimmerTime(M)
End If
Close #13

N = 0
Ctr = 0
Found = False
Open App.Path & "\1650.txt" For Input As #14
Do Until EOF(14)
    Ctr = Ctr + 1
    Input #14, Swimmer(Ctr), SwimmerTime(Ctr)
Loop
Do While ((Not Found) And (N < Ctr))
    N = N + 1
    If SwimmerInput = Swimmer(N) Then Found = True
Loop

If Found Then
    picResults.Print SwimmerInput; " swam the 1650 Freestyle with a time of "; SwimmerTime(N)
End If
Close #14

If Not Found Then
    response = MsgBox("Name not recognized; check spelling or capitalization.", , "Error")
End If
End Sub

