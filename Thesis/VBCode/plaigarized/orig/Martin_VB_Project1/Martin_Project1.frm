VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18270
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   18270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdform3 
      Caption         =   "Refine Search Results"
      Height          =   1455
      Left            =   960
      TabIndex        =   15
      Top             =   10560
      Width           =   2295
   End
   Begin VB.CommandButton cmdform2 
      Caption         =   "Refine Display Results"
      Height          =   1455
      Left            =   960
      TabIndex        =   13
      Top             =   8760
      Width           =   2295
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Exit Program "
      Height          =   1575
      Left            =   11760
      TabIndex        =   12
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Previous Entry "
      Height          =   1575
      Left            =   11760
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      Height          =   7695
      Left            =   3960
      ScaleHeight     =   7635
      ScaleWidth      =   6915
      TabIndex        =   7
      Top             =   2520
      Width           =   6975
   End
   Begin VB.CommandButton cmdmetacritic 
      Caption         =   "View metacritic.com Top 50 Films of All Time"
      Height          =   1575
      Left            =   960
      TabIndex        =   6
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdafi 
      Caption         =   "View AFI.com Top 50 Films of All Time"
      Height          =   1575
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdimdb 
      Caption         =   "View IMDB.com Top 50 Films of All Time"
      Height          =   1575
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmddisclaimer 
      Caption         =   "Disclaimer"
      Height          =   1335
      Left            =   14400
      TabIndex        =   2
      Top             =   8760
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Click Here To:"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Click Here To:"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   9240
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Click Here To:"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Click Here To:"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Click Here To:"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Please click here=======> To read a letter from the Editor to pertaining to the discrepancies of this program results....  "
      Height          =   975
      Left            =   11880
      TabIndex        =   3
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   $"Martin_Project1.frx":0000
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   $"Martin_Project1.frx":00ED
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdafi_Click()
Dim l As Integer


CTR3 = 0
Open App.Path & "\afilist.txt" For Input As #1
Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR3 = CTR3 + 1
    Input #1, num(CTR3), tit(CTR3), dir(CTR3)
    Loop
picresults.Print "Number"; Tab(20); "Title"; Tab(40); "Director"
For l = 1 To CTR3
        picresults.Print num(l); Tab(20); tit(l); Tab(40); dir(l)
Next l
Close #1
'disable the command button for printing above
    cmdafi.Enabled = False
End Sub

Private Sub cmdclear_Click()
picresults.Cls

End Sub

Private Sub cmddisclaimer_Click()
picresults.Print "The results of the program does not create a consensus amongst the best 50 films of all time,"
picresults.Print "it is merely a tool to compare and contrast the three most popular lists."
picresults.Print "This is due to the fact that all three charts are biased in some nature."
picresults.Print "For instance, movies on IMDB.com are rated by people from around the world"
picresults.Print "who choose to sign up as a free member of the website."
picresults.Print "In addition, the AFI.com chart is determined by a group of elected board members"
picresults.Print "who represent the American film industry. Furthermore, the chart given is from the year 2008."
picresults.Print "Finally, Metacritic.com creates an average score of rates that are given to movies"
picresults.Print "by professional critics who represent major publications."
picresults.Print "This website primarily focuses on movies reviewed in the past decade."
End Sub

Private Sub cmdform2_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub cmdform3_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub cmdimdb_Click()
Dim J As Integer


CTR = 0
Open App.Path & "\imdblist.txt" For Input As #1
Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
    Input #1, number(CTR), title(CTR), rating(CTR), director(CTR)
    Loop
picresults.Print "number"; Tab(20); "title"; Tab(40); "rating"; Tab(60); "director"
For J = 1 To CTR

        picresults.Print number(J); Tab(20); title(J); Tab(40); rating(J); Tab(60); director(J)
Next J
Close #1
'disable the command button for printing above
    cmdimdb.Enabled = False
End Sub

Private Sub cmdmetacritic_Click()
Dim n As Integer

CTR2 = 0
Open App.Path & "\metacritic.txt" For Input As #1
Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1
    Input #1, numbers(CTR2), titles(CTR2), ratings(CTR2), directors(CTR2)
    Loop
picresults.Print "Number"; Tab(20); "Title"; Tab(40); "Rating"; Tab(60); "Director"
For n = 1 To CTR2
        picresults.Print numbers(n); Tab(20); titles(n); Tab(40); ratings(n); Tab(60); directors(n)
Next n
Close #1
'disable the command button for printing above
    cmdmetacritic.Enabled = False

End Sub

Private Sub cmdquit_Click()

End

End Sub

