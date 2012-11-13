VERSION 5.00
Begin VB.Form frmSudoku 
   BackColor       =   &H00000000&
   Caption         =   "Sudoku"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHighScores 
      Caption         =   "High Scores"
      Height          =   495
      Left            =   3840
      TabIndex        =   89
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdWorksCited 
      Caption         =   "Works Cited"
      Height          =   495
      Left            =   3840
      TabIndex        =   88
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History"
      Height          =   495
      Left            =   3840
      TabIndex        =   87
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdname 
      Caption         =   "Enter Name"
      Height          =   495
      Left            =   2520
      TabIndex        =   86
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txt21 
      Height          =   285
      Left            =   1320
      TabIndex        =   84
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text78 
      Height          =   285
      Left            =   3240
      TabIndex        =   83
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text77 
      Height          =   285
      Left            =   2880
      TabIndex        =   82
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text76 
      Height          =   285
      Left            =   3240
      TabIndex        =   81
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text75 
      Height          =   285
      Left            =   2880
      TabIndex        =   80
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text74 
      Height          =   285
      Left            =   2520
      TabIndex        =   79
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text73 
      Height          =   285
      Left            =   2520
      TabIndex        =   78
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text72 
      Height          =   285
      Left            =   3240
      TabIndex        =   77
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text71 
      Height          =   285
      Left            =   2880
      TabIndex        =   76
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text70 
      Height          =   285
      Left            =   2520
      TabIndex        =   75
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text69 
      Height          =   285
      Left            =   3240
      TabIndex        =   74
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text68 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   73
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text67 
      Height          =   285
      Left            =   3240
      TabIndex        =   72
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text66 
      Height          =   285
      Left            =   2880
      TabIndex        =   71
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text65 
      Height          =   285
      Left            =   3240
      TabIndex        =   70
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text64 
      Height          =   285
      Left            =   2880
      TabIndex        =   69
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text63 
      Height          =   285
      Left            =   2520
      TabIndex        =   68
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text62 
      Height          =   285
      Left            =   2520
      TabIndex        =   67
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text61 
      Height          =   285
      Left            =   2520
      TabIndex        =   66
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text60 
      Height          =   285
      Left            =   2040
      TabIndex        =   65
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text59 
      Height          =   285
      Left            =   1680
      TabIndex        =   64
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text58 
      Height          =   285
      Left            =   1320
      TabIndex        =   63
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text57 
      Height          =   285
      Left            =   2040
      TabIndex        =   62
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text56 
      Height          =   285
      Left            =   1680
      TabIndex        =   61
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text55 
      Height          =   285
      Left            =   1320
      TabIndex        =   60
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text54 
      Height          =   285
      Left            =   2040
      TabIndex        =   59
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text53 
      Height          =   285
      Left            =   1680
      TabIndex        =   58
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text52 
      Height          =   285
      Left            =   1320
      TabIndex        =   57
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text51 
      Height          =   285
      Left            =   2040
      TabIndex        =   56
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text50 
      Height          =   285
      Left            =   1680
      TabIndex        =   55
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text49 
      Height          =   285
      Left            =   1320
      TabIndex        =   54
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text48 
      Height          =   285
      Left            =   2040
      TabIndex        =   53
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text47 
      Height          =   285
      Left            =   1680
      TabIndex        =   52
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text46 
      Height          =   285
      Left            =   1320
      TabIndex        =   51
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text45 
      Height          =   285
      Left            =   2040
      TabIndex        =   50
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text44 
      Height          =   285
      Left            =   1680
      TabIndex        =   49
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text43 
      Height          =   285
      Left            =   1320
      TabIndex        =   48
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text42 
      Height          =   285
      Left            =   840
      TabIndex        =   47
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text41 
      Height          =   285
      Left            =   480
      TabIndex        =   46
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text40 
      Height          =   285
      Left            =   840
      TabIndex        =   45
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text39 
      Height          =   285
      Left            =   480
      TabIndex        =   44
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text38 
      Height          =   285
      Left            =   840
      TabIndex        =   43
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text37 
      Height          =   285
      Left            =   480
      TabIndex        =   42
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text36 
      Height          =   285
      Left            =   840
      TabIndex        =   41
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text35 
      Height          =   285
      Left            =   480
      TabIndex        =   40
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text34 
      Height          =   285
      Left            =   840
      TabIndex        =   39
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text33 
      Height          =   285
      Left            =   480
      TabIndex        =   38
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text32 
      Height          =   285
      Left            =   840
      TabIndex        =   37
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text31 
      Height          =   285
      Left            =   480
      TabIndex        =   36
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text30 
      Height          =   285
      Left            =   3240
      TabIndex        =   35
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   2880
      TabIndex        =   34
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   2520
      TabIndex        =   33
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt33 
      Height          =   285
      Left            =   3240
      TabIndex        =   32
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt32 
      Height          =   285
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt31 
      Height          =   285
      Left            =   2520
      TabIndex        =   30
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt26 
      Height          =   285
      Left            =   2040
      TabIndex        =   29
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt25 
      Height          =   285
      Left            =   1680
      TabIndex        =   28
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt23 
      Height          =   285
      Left            =   2040
      TabIndex        =   27
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt22 
      Height          =   285
      Left            =   1680
      TabIndex        =   26
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt24 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3240
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   2520
      TabIndex        =   22
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt29 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt28 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt27 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdHard 
      Caption         =   "Hard"
      Height          =   495
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdMedium 
      Caption         =   "Medium"
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdEasy 
      BackColor       =   &H0000FF00&
      Caption         =   "Easy"
      Height          =   495
      Index           =   0
      Left            =   3840
      MaskColor       =   &H0000FF00&
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txt12 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt13 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt14 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt15 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt16 
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txt17 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt18 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt19 
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txt11 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   4920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblSudoku 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   -1560
      TabIndex        =   85
      Top             =   3720
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   3480
   End
End
Attribute VB_Name = "frmSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is used to play Sudoku. It has three different difficulties to choose from along with high scores, history, works cited, and enter name.
Option Explicit

'This button brings you to the hard puzzle
Private Sub cmdHard_Click(Index As Integer)
    frmSudoku.Hide
    frmHard.Show
End Sub

'This button brings you to the High Scores
Private Sub cmdHighScores_Click()
    frmSudoku.Hide
    frmHighScores.Show
End Sub

'This button shows you the history
Private Sub cmdHistory_Click()
    MsgBox "Sudoku(soo-doh-koo) is a logic-based, combinatorial number-placement puzzle. The objective is to fill a 9×9 grid with digits so that each column, each row, and each of the nine 3×3 sub-grids contain all of the digits from 1 to 9. The puzzle setter provides a partially completed grid, which  has a unique solution. The puzzle was popularized in 1986 by the Japanese puzzle company Nikoli, under the name Sudoku, meaning single number. It became an international hit in 2005."
End Sub
'This button brings you to the medium puzzle
Private Sub cmdMedium_Click(Index As Integer)
    frmSudoku.Hide
    frmMedium.Show
End Sub

'This button asks to input your first and last name
Private Sub cmdname_Click()
    Dim spaceLocation As Integer
    
    EnterName = InputBox("Hello, please enter your first and last name. (seperated by a space)")
    MsgBox "Hello, " & EnterName & " welcome to the world of Sudoku!", , "Welcome"
    
    spaceLocation = InStr(EnterName, " ")
    FirstName = Left(EnterName, spaceLocation - 1)
    LastName = Right(EnterName, Len(EnterName) - spaceLocation)
End Sub

'This button quits the program
Private Sub cmdQuit_Click()
    End
End Sub

'This button brings you to the easy puzzle
Private Sub cmdEasy_Click(Index As Integer)
    frmSudoku.Hide
    frmEasy.Show
End Sub

'This button brings you to the works cited
Private Sub cmdWorksCited_Click()
    frmWorksCited.Show
End Sub

'This messages appears upon starting the program
Private Sub Form_Load()
    MsgBox "Hello, welcome to the wonderful world of Sudoku. Please enter your name so we can personalize your experience here.", , "Welcome"
End Sub
