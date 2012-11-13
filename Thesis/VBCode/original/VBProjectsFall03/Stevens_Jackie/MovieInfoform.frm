VERSION 5.00
Begin VB.Form MovieInfo 
   BackColor       =   &H00800080&
   Caption         =   "Movies: Movie Info"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FF0000&
      Caption         =   "Look Up Movie Info"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtMovieNumber 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label18 
      BackColor       =   &H00800080&
      Caption         =   "Enter the number of the movie you wish to access:"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   7815
   End
   Begin VB.Label Label17 
      BackColor       =   &H00800080&
      Caption         =   "17. Under the Tuscan Sun"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "16. The Texas Chainsaw         Massacre"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6000
      TabIndex        =   15
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label15 
      BackColor       =   &H00800080&
      Caption         =   "15. Secondhand Lions"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label14 
      BackColor       =   &H00800080&
      Caption         =   "14. School of Rock "
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00800080&
      Caption         =   "13. The Rundown"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00800080&
      Caption         =   "12. Runaway Jury"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00800080&
      Caption         =   "11. Pirates of the                Caribbean"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00800080&
      Caption         =   "10. Out of Time"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H00800080&
      Caption         =   "9. Once Upon a Time     in Mexico"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800080&
      Caption         =   "8. Mystic River"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800080&
      Caption         =   "7. Matchstick Men"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800080&
      Caption         =   "6. Lost in Translation"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Caption         =   "5. Kill Bill: Volume 1"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800080&
      Caption         =   "4. Intolerable Cruelty"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800080&
      Caption         =   "3. House of the Dead"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "2. Good Boy"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "1. Cold Creek Manor   "
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "MovieInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: MovieInfo (MovieInfoform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose: 1. To display numbers for movie titles that will allow the user
            'to look up information about a particular movie and take them
            'to that form
         '2. To provide links to move to other forms in the program.

Option Explicit

Private Sub cmdBack_Click()
    'go back to main movie page
MovieMain.Show
MovieInfo.Hide
End Sub

Private Sub cmdFind_Click()
    'declare local variables
Dim Selection As Single

Selection = Val(txtMovieNumber.Text)
    'compares selection number and finds info, displays info
   
Do While Selection < 1 Or Selection > 17
    MsgBox "Please enter a number between 1 and 17.", , "Invalid Entry"
    Selection = InputBox("Enter a number between 1 and 17", "Movie Selection")
Loop
        
If Selection = 1 Then
        MovieInfo.Hide
        ColdCreekManor.Show
    ElseIf Selection = 2 Then
        MovieInfo.Hide
        GoodBoy.Show
    ElseIf Selection = 3 Then
        MovieInfo.Hide
        HouseoftheDead.Show
    ElseIf Selection = 4 Then
        MovieInfo.Hide
        IntolerableCruelty.Show
    ElseIf Selection = 5 Then
        MovieInfo.Hide
        KillBill.Show
    ElseIf Selection = 6 Then
        MovieInfo.Hide
        LostinTranslation.Show
    ElseIf Selection = 7 Then
        MovieInfo.Hide
        MatchstickMen.Show
    ElseIf Selection = 8 Then
        MovieInfo.Hide
        MysticRiver.Show
    ElseIf Selection = 9 Then
        MovieInfo.Hide
        OnceUponaTimeinMexico.Show
    ElseIf Selection = 10 Then
        MovieInfo.Hide
        OutOfTime.Show
    ElseIf Selection = 11 Then
        MovieInfo.Hide
        PiratesoftheCaribbean.Show
    ElseIf Selection = 12 Then
        MovieInfo.Hide
        RunawayJury.Show
    ElseIf Selection = 13 Then
        MovieInfo.Hide
        Rundown.Show
    ElseIf Selection = 14 Then
        MovieInfo.Hide
        SchoolOfRock.Show
    ElseIf Selection = 15 Then
        MovieInfo.Hide
        SecondhandLions.Show
    ElseIf Selection = 16 Then
        MovieInfo.Hide
        TexasChainsawMassacre.Show
    ElseIf Selection = 17 Then
        MovieInfo.Hide
        UndertheTuscanSun.Show
    ElseIf Selection < 1 Then
        MsgBox "You must enter a number between 1 and 17", , "Invalid Entry"
    ElseIf Selection > 17 Then
        MsgBox "You must enter a number between 1 and 17", , "Invalid Entry"
End If


End Sub

Private Sub cmdQuit_Click()
    'Quits program
End
End Sub

