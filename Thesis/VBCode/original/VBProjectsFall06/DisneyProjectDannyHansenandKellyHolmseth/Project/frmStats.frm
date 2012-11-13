VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00004080&
   Caption         =   "Sorting"
   ClientHeight    =   6705
   ClientLeft      =   3540
   ClientTop       =   2760
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8445
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C0C0&
      Height          =   1095
      Left            =   2880
      ScaleHeight     =   1035
      ScaleWidth      =   5355
      TabIndex        =   5
      Top             =   5160
      Width           =   5415
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H0000C0C0&
      Caption         =   "Average Rating"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   2640
      Picture         =   "frmStats.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton cmdRating 
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort by Ratings On a Scale of 1-10"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort by Alphabetical Order"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdRelease 
      BackColor       =   &H0000C0C0&
      Caption         =   "Sort by Release Date"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/31/06
'Objective: The objective of this form is to show the users different options of ways they can view movie data
Option Explicit
Dim I As Integer
Dim Pass As Integer
Dim TempDate As Date
Dim TempName As String
Dim TempRating As Single

Private Sub cmdAverage_Click()
Dim Average As Single
Dim I As Integer
For I = 1 To 10     'ten different movies- set array
Average = Average + MovieRating(I)      'averages the movie ratings
Next I
picResults.Print Average / 10, "Is the average rating for these 10 wonderful movies"
End Sub

Private Sub cmdBack_Click()
frmStats.Hide       'allows user to go to the Top form
frmTop.Show
End Sub

Private Sub cmdName_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If MovieName(I) > MovieName(I + 1) Then 'sorting
            TempName = MovieName(I)
            MovieName(I) = MovieName(I + 1)
            MovieName(I + 1) = TempName
            TempDate = MovieRelease(I)
            MovieRelease(I) = MovieRelease(I + 1)
            MovieRelease(I + 1) = TempDate
            TempRating = MovieRating(I)
            MovieRating(I) = MovieRating(I + 1)
            MovieRating(I + 1) = TempRating
            
        End If
    Next I
Next Pass
frmDisplay.showStats
frmDisplay.Show     'allows user to go to the display form
frmStats.Hide
End Sub

Private Sub cmdRating_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If MovieRating(I) > MovieRating(I + 1) Then     'sorting
            TempRating = MovieRating(I)
            MovieRating(I) = MovieRating(I + 1)
            MovieRating(I + 1) = TempRating
            TempDate = MovieRelease(I)
            MovieRelease(I) = MovieRelease(I + 1)
            MovieRelease(I + 1) = TempDate
            TempName = MovieName(I)
            MovieName(I) = MovieName(I + 1)
            MovieName(I + 1) = TempName
            
        End If
    Next I
Next Pass
frmDisplay.Show     'allows user to go to the display form
frmStats.Hide
frmDisplay.showStats
End Sub

Private Sub cmdRelease_Click()
For Pass = 1 To 9
    For I = 1 To 10 - Pass
        If MovieRelease(I) > MovieRelease(I + 1) Then 'Bubble Sort
            TempDate = MovieRelease(I)  'bubble sort takes advantage of the temp feature.  As it goes down through the list it compares and switches when need be based on if it is sorting by rating, alphabet, or release date.  It makes n-1 comparisons at the most so in this case 10-1 = 9 is the most number of passes this bubble sort will do.
            MovieRelease(I) = MovieRelease(I + 1)
            MovieRelease(I + 1) = TempDate
            TempName = MovieName(I)
            MovieName(I) = MovieName(I + 1)
            MovieName(I + 1) = TempName
            TempRating = MovieRating(I)
            MovieRating(I) = MovieRating(I + 1)
            MovieRating(I + 1) = TempRating
            
        End If
    Next I
Next Pass
frmDisplay.showStats        'Allows user to go to the display form
frmDisplay.Show
frmStats.Hide
End Sub


Private Sub Form_Load()
Counter = 0
Open App.Path & "\Stats.txt" For Input As #1
Do Until EOF(1) ' tells computer to read from array until the end of the file is reached.
    Counter = Counter + 1       'Counter for the loop
    Input #1, MovieName(Counter), MovieRelease(Counter), MovieRating(Counter)
Loop
Close #1 'it is very important to close #1 otherwise your program won't work
End Sub
