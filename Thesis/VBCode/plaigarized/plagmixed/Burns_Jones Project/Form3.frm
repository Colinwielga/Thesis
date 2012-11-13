VERSION 5.00
Begin VB.Form Form3
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   12765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18195
   LinkTopic       =   "Form3"
   ScaleHeight     =   12765
   ScaleWidth      =   18195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSerena
      Caption         =   "Serena Williams"
      Height          =   1215
      Left            =   19080
      TabIndex        =   12
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdKim
      Caption         =   "Kim Clijster"
      Height          =   1215
      Left            =   19080
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdJuan
      Caption         =   "Juan Martin Del Potro"
      Height          =   1335
      Left            =   19080
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdRoger
      Caption         =   "Roger Federer"
      Height          =   1455
      Left            =   19080
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdYearWomen
      Caption         =   "Enter a year between 2004 and 2009 and I will tell you the winner on the women's side of the draw!"
      BeginProperty Font
         Name            =   "Footlight MT Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14040
      TabIndex        =   7
      Top             =   9720
      Width           =   3615
   End
   Begin VB.CommandButton cmdSearchYear
      Caption         =   "Enter a year between 2004 and 2009 and I will tell you the winner on the men's side!"
      BeginProperty Font
         Name            =   "Footlight MT Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   14160
      TabIndex        =   6
      Top             =   7680
      Width           =   3615
   End
   Begin VB.CommandButton cmdWomensWinners
      Caption         =   "I wonder who won the women's titles?"
      BeginProperty Font
         Name            =   "Footlight MT Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   14040
      TabIndex        =   3
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CommandButton cmdMensWinners
      Caption         =   "I wonder who won the men's titles?"
      BeginProperty Font
         Name            =   "Footlight MT Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14160
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton cmdGoBackToForum
      Caption         =   "Return To Main Page"
      BeginProperty Font
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   13920
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.PictureBox Picture1
      Height          =   12975
      Left            =   840
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   12915
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.PictureBox picResults
         Height          =   5775
         Left            =   0
         ScaleHeight     =   5715
         ScaleWidth      =   5835
         TabIndex        =   9
         Top             =   4200
         Width           =   5895
      End
      Begin VB.PictureBox picWomensWinners
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   6675
         TabIndex        =   5
         Top             =   1800
         Width           =   6735
      End
      Begin VB.PictureBox picMensWinners
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   6675
         TabIndex        =   4
         Top             =   0
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGoBackToForum_Click()
Form3.Hide
Form2.Show
'this takes the user back to the main menu'
End Sub

Private Sub cmdJuan_Click()
picResults.Picture = LoadPicture(App.Path & "\Juan.jpg")
'loads picture onto screen'
' lskadfj dslfk sdlk fsdlfksd
End Sub


Private Sub cmdKim_Click()
picResults.Picture = LoadPicture(App.Path & "\Kim.jpg")
'loads picture onto screen'
End Sub

' lskadfj dslfk sdlk fsdlfksd
Private Sub cmdMensWinners_Click()
Dim I As Integer
Dim JJJJ As String

Open App.Path & "\USopenmens.txt" For Input As #1
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, MFirstName(Ctr), MLastName(Ctr), MYear(Ctr)
Loop
Close #1

MsgBox "The Winners Have Been Entered!"
picMensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picMensWinners.Print "***********************************"
For I = 1 To Ctr
picMensWinners.Print MLastName(I), MFirstName(I); Tab(35); MYear(I)
Next I

'this button helps to get the information from a text file on notepad and to display it on the screen'


End Sub

Private Sub cmdRoger_Click()
picResults.Picture = LoadPicture(App.Path & "\Roger.jpg")
'loads picture onto screen'
End Sub
' lskadfj dslfk sdlk fsdlfksd

Private Sub cmdSearchYear_Click()
Dim Found As Boolean, I As Integer, Years As String
Years = InputBox("Enter A year And I Will Tell You The Winner!")
I = 0
Found = False
Do While ((Not Found) And (I < Ctr))
I = I
I = I + 1
If Years = MYear(I) Then
Found = True
End If
Loop
If Found <> True Then
MsgBox Years & "no winner entered this year"
ElseIf True
MsgBox Yearsn & "The player " & MFirstName(I) & " " & MLastName(I) & " Was the winner this Year."
End If

'this button helps to search through the array to determine what player won the championship in what year'



End Sub

Private Sub cmdSerena_Click()
picResults.Picture = LoadPicture(App.Path & "\Serena.jpg")
'loads picture onto screen'
End Sub

Private Sub cmdWomensWinners_Click()
Open App.Path & "\USopenwomens.txt" For Input As #1
' lskadfj dslfk sdlk fsdlfksd
Ctr = 0
Do While Not EOF(1)
Ctr = Ctr + 1
Input #1, WFirstName(Ctr), WLastName(Ctr), WYear(Ctr)
Loop
Close #1
MsgBox "The Winners Have Been Entered!"
' lskadfj dslfk sdlk fsdlfksd
picWomensWinners.Print "Last Name", "First Name"; Tab(35); "Year"
picWomensWinners.Print "***************************************"
For I = 1 To Ctr
picWomensWinners.Print WLastName(I), WFirstName(I); Tab(35); WYear(I)
Next I

'this button helps to get the information from a text file on notepad and to display it on the screen'

End Sub

Private Sub cmdYearWomen_Click()
Dim Found As Boolean, I As Integer, Years As String
Years = InputBox("Enter A year And I Will Tell You The Winner!")
I = 0
' lskadfj dslfk sdlk fsdlfksd
Found = False
Do While ((Not Found) And (I < Ctr))
I = I + 1
If Years = WYear(I) Then
Found = True
End If
Loop
If (Not Found) Then
' lskadfj dslfk sdlk fsdlfksd
MsgBox Years & "no winner entered this year"
Else
MsgBox Yearsn & "The player " & WFirstName(I) & " " & WLastName(I) & " Was the winner this Year."
End If
End Sub
'this button helps to search through the array to determine what player won the championship in what year'
