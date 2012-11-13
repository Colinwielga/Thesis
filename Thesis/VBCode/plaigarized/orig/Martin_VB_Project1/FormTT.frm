VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000013&
   Caption         =   "Form3"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   LinkTopic       =   "Form3"
   ScaleHeight     =   9810
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Exit Program"
      Height          =   1575
      Left            =   10800
      TabIndex        =   8
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Homepage"
      Height          =   1575
      Left            =   10800
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Previous Entry"
      Height          =   1575
      Left            =   10800
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdrating 
      Caption         =   "Search a Rating Score to View All the Movies that have Recieved the entered Score"
      Height          =   1575
      Left            =   10920
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmddirector 
      Caption         =   "Search the Disired Director to View All of His/Her Films that Rank Amongst the Top Movie Charts"
      Height          =   1335
      Left            =   7680
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "Search the Desired  Rating Place of 1-50 to Compare the Place Amongst the Three Charts"
      Height          =   1335
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdtitle 
      Caption         =   "Enter the Title For Your Movie of Choice"
      Height          =   1335
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.PictureBox picresults 
      Height          =   6015
      Left            =   4200
      ScaleHeight     =   5955
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   3240
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   $"Form3.frx":0000
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdclear_Click()
picresults.Cls

End Sub

Private Sub cmddirector_Click()
Dim b As Integer, movie As String, Found As Boolean
b = 0
movie = InputBox("Enter the name of a Movie Director", "director")
Do While ((Not Found) And (b < CTR))
b = b + 1
If movie = director(b) Or movie = directors(b) Or movie = dir(b) Then
    Found = True
    End If
    Loop
    
    If (Not Found) Then
        picresults.Print "This movie Director does not have any films in the top 50 charts"
        Else
        picresults.Print "The director  "; director(b); "  has these films ranked in the top 50 charts"
    End If
End Sub

Private Sub cmdnumber_Click()
Dim a As Integer, movie As String, Found As Boolean

a = 0
movie = InputBox("Enter the number rank within the top 50", "number")
Do While ((Not Found) And (a < CTR))
a = a + 1

If movie = number(a) Or movie = numbers(a) Or movie = num(a) Then
    Found = True
    End If
    Loop
    
    If (Not Found) Then
        picresults.Print movie; "  The charts do not consider this placements value"
        Else
        picresults.Print movie; "  There are movies that are ranked in these charts at this placement"
    End If
End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdrating_Click()
Dim d As Integer, Found As Boolean, movie As String

d = 0
movie = InputBox("Enter the Rating of a Movie", "rating")
Do While ((Not Found) And (d < CTR))
d = d + 1
If movie = rating(d) Or movie = ratings(d) Then
    Found = True
    End If
    Loop
    
    If (Not Found) Then
        picresults.Print " No movies were bad enough to recieve this rating in the top 50 charts"
        Else
        picresults.Print "The rating of  "; movie; "  has been rewarded to these Movies"
    End If

End Sub

Private Sub cmdreturn_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub cmdtitle_Click()
Dim c As Integer, movie As String, Found As Boolean

c = 0
movie = InputBox("Enter the Title of a Movie", "movie")
Do While ((Not Found) And (c < CTR))
c = c + 1
If movie = title(c) Or movie = titles(c) Or movie = tit(c) Then
    Found = True
    End If
    Loop
    
    If (Not Found) Then
        picresults.Print movie; "  This Movie is not in any of the Top 50 charts"
        Else
        picresults.Print movie; "  is listed in these charts"
    End If

End Sub

