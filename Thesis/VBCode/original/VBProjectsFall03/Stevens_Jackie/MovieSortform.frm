VERSION 5.00
Begin VB.Form MovieSort 
   BackColor       =   &H00800080&
   Caption         =   "Sort Movies"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5655
      Left            =   2880
      ScaleHeight     =   5595
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   480
      Width           =   7455
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort Movies Alphabetically"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdRating 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort Movies by Rating"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "MovieSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: MovieSort (MovieSortform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose: 1. To display the movies in a particular order so the user can
            'find a movie they are interested more quickly by rating or by
            'alphabetical title
         '2. To provide links to move to other forms in the program.


Option Explicit
    'declare global variables
Dim TmpScreen As String, TmpMovie As String, TmpRating As String
Dim Pass As Integer, Comp As Integer
    
Private Sub cmdAlpha_Click()
    'clear screen
picResults.Cls
    'initialize counter
CTR = 0
Open Path & "MovieFile.txt" For Input As #1
    'load array
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Screen(CTR), Movie(CTR), Rating(CTR)
Loop
Close
    'Sort alphabetically by title

For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Movie(Comp) > Movie(Comp + 1) Then
            TmpRating = Rating(Comp)
            Rating(Comp) = Rating(Comp + 1)
            Rating(Comp + 1) = TmpRating
            TmpMovie = Movie(Comp)
            Movie(Comp) = Movie(Comp + 1)
            Movie(Comp + 1) = TmpMovie
            TmpScreen = Screen(Comp)
            Screen(Comp) = Screen(Comp + 1)
            Screen(Comp + 1) = TmpScreen
        End If
    Next Comp
Next Pass
    'print sorted list
For J = 1 To 18
    picResults.Print Movie(J); Tab(62); Rating(J)
Next J
End Sub

Private Sub cmdBack_Click()
    'Go back to main movie form
MovieMain.Show
MovieSort.Hide
End Sub

Private Sub cmdRating_Click()
    'clear screen
picResults.Cls
    'initialize counter
CTR = 0
Open Path & "MovieFile.txt" For Input As #1
    'load array
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Screen(CTR), Movie(CTR), Rating(CTR)
Loop
Close
    'sort by rating
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
        If Rating(Comp) > Rating(Comp + 1) Then
            TmpRating = Rating(Comp)
            Rating(Comp) = Rating(Comp + 1)
            Rating(Comp + 1) = TmpRating
            TmpMovie = Movie(Comp)
            Movie(Comp) = Movie(Comp + 1)
            Movie(Comp + 1) = TmpMovie
            TmpScreen = Screen(Comp)
            Screen(Comp) = Screen(Comp + 1)
            Screen(Comp + 1) = TmpScreen
        End If
    Next Comp
Next Pass
    'print sorted list
For J = 1 To 18
    picResults.Print Rating(J); Tab(12); Movie(J)
Next J
End Sub

