VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   Caption         =   "Form2"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form2"
   ScaleHeight     =   10125
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd50 
      Caption         =   "Add or View the Top 40 to 50"
      Height          =   1095
      Left            =   1800
      TabIndex        =   10
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmd40 
      Caption         =   "Add or View the Top 30 to 40"
      Height          =   1095
      Left            =   1800
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmd30 
      Caption         =   "Add or View the Top 20 to 30"
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmd20 
      Caption         =   "Add or View the Top 10 to 20"
      Height          =   1095
      Left            =   1800
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmd10 
      Caption         =   "View the Top Ten Forms All Charts"
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdalphabetical 
      Caption         =   "View All Three Charts in Alphabetical Order"
      Height          =   1095
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Previous Entry"
      Height          =   1575
      Left            =   11280
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Exit Program"
      Height          =   1575
      Left            =   11280
      TabIndex        =   3
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H80000015&
      Caption         =   "Return to Homepage"
      Height          =   1575
      Left            =   11280
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picresults 
      Height          =   8655
      Left            =   4440
      ScaleHeight     =   8595
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "If desired, one is able to view all three Top Fifty Movie Charts in additional formats that are listed below."
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd10_Click()
'THIS IS TO SHOW RANK 1-10 ON ALL THREE ARRAYS
Dim l As Integer
For l = 1 To CTR
If number(l) <= 10 Then
'MUST PUT ALL FORMS OF SAME FILE TYPE TO PRINT ALL TO COMPARE
    picresults.Print "the Movie  "; title(l); Tab(20); number(l); "  is in the Top Ten."
If numbers(l) <= 10 Then
    picresults.Print "The Movie  "; titles(l); Tab(20); numbers(l); "  is in the Top Ten."
If num(l) <= 10 Then
    picresults.Print "The Movie  "; tit(l); Tab(20); num(l); "  is in the Top Ten."
Else
    picresults.Print "Error"
End If
End If
End If
Next l
End Sub

Private Sub cmd20_Click()
'DIM h as the new variable for use
'must change variable letter for every new subroutine
Dim h As Integer
For h = 1 To CTR
If number(h) > 10 And number(h) <= 20 Then
'showing the next ten films listed on all charts
    picresults.Print "The Movie  "; title(h); Tab(20); number(h); "  is in the Top Twenty."
If numbers(h) > 10 And numbers(h) <= 20 Then
    picresults.Print "The  "; titles(h); Tab(20); numbers(h); "  is in the Top Twenty."
If num(h) > 10 And num(h) <= 20 Then
    picresults.Print "The  "; tit(h); Tab(20); num(h); "  is in the Top Twenty."
Else
    picresults.Print "Error"
End If
End If
End If
Next h
End Sub

Private Sub cmd30_Click()
Dim g As Integer
For g = 1 To CTR
If number(g) > 20 And number(g) <= 30 Then
'same at other buttons on form
'adds next ten movies 21-30
    picresults.Print "The  "; title(g); Tab(20); number(g); "  is in the Top Thirty."
If numbers(g) > 20 And numbers(g) <= 30 Then
    picresults.Print "The  "; titles(g); Tab(20); numbers(g); "  is in the Top Thirty."
If num(g) > 20 And num(g) <= 30 Then
    picresults.Print "The  "; tit(g); Tab(20); num(g); "  is in the Top Thirty."
Else
    picresults.Print "Error"
End If
End If
End If
Next g
End Sub

Private Sub cmd40_Click()
Dim R As Integer
For R = 1 To CTR
If number(R) > 30 And number(R) <= 40 Then
    picresults.Print "The  "; title(R); Tab(20); number(R); "  is in the Top Forty."
If numbers(R) > 30 And numbers(R) <= 40 Then
    picresults.Print "The  "; titles(R); Tab(20); numbers(R); "  is in the Top Forty."
If num(R) > 30 And num(R) <= 40 Then
    picresults.Print "The  "; tit(R); Tab(20); num(R); "  is in the Top Forty."
Else
    picresults.Print "Error"
End If ' must have 3 end ifs because theres 3 different arrays to go through
End If
End If
Next R
End Sub

Private Sub cmd50_Click()
Dim w As Integer
For w = 1 To CTR
If number(w) > 40 And number(w) <= 50 Then
'end of each text file
    picresults.Print "The  "; title(w); Tab(20); number(w); "  is in the Top Fifty."
If numbers(w) > 40 And numbers(w) <= 50 Then
    picresults.Print "The  "; titles(w); Tab(20); numbers(w); "  is in the Top Fifty."
If num(w) > 40 And num(w) <= 50 Then
    picresults.Print "The  "; tit(w); Tab(20); num(w); "  is in the Top Fifty."
Else
    picresults.Print "Error"
End If
End If
End If
Next w
End Sub

Private Sub cmdalphabetical_Click()
Dim J As Integer, Pass As Integer, Pos As Integer, Temp1 As String, Temp2 As Integer, Temp3 As String, Temp4 As Integer, Temp5 As String, temp6 As Integer
' for alphabetical order of movies
'must have 6 temps for two columns of the three arrays
' listing only the titles and the rank/ position in the chart
For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If title(Pos) > title(Pos + 1) Then
            Temp1 = title(Pos)
            title(Pos) = title(Pos + 1)
            title(Pos + 1) = Temp1
                Temp2 = number(Pos)
                number(Pos) = number(Pos + 1)
                number(Pos + 1) = Temp2
            Temp3 = titles(Pos)
            titles(Pos) = titles(Pos + 1)
            titles(Pos + 1) = Temp3
                Temp4 = numbers(Pos)
                numbers(Pos) = numbers(Pos + 1)
                numbers(Pos + 1) = Temp4
            Temp5 = tit(Pos)
            tit(Pos) = tit(Pos + 1)
            tit(Pos + 1) = Temp5
                temp6 = num(Pos)
                num(Pos) = num(Pos + 1)
                num(Pos + 1) = temp6
    End If
Next Pos
Next Pass

'use tab 20 to keep all 3 listsin line of one another
For J = 1 To CTR
picresults.Print title(J); Tab(20); number(J)
picresults.Print titles(J); Tab(20); numbers(J)
picresults.Print tit(J); Tab(20); num(J)
Next J
End Sub

Private Sub cmdclear_Click()
picresults.Cls
'clears the picturebox
End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdreturn_Click()
Form1.Show
Form2.Hide
'go back to the 1st form
End Sub
