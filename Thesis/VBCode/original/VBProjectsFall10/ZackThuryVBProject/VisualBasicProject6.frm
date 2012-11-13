VERSION 5.00
Begin VB.Form FrmTotalRecords 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReadYourRecords 
      Caption         =   "Show All of Your Records"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdBackToYourRecords 
      Caption         =   "Add a New Record"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picYourTotalRecords 
      BackColor       =   &H00FFFFFF&
      Height          =   9255
      Left            =   2880
      ScaleHeight     =   9195
      ScaleWidth      =   9075
      TabIndex        =   3
      Top             =   1560
      Width           =   9135
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Your Average Score Per 9 Holes"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdBackToHome 
      BackColor       =   &H00808000&
      Caption         =   "Back To Home Screen"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblYourRecords 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Records"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmTotalRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form opens a file of your records and prints them for you to see and calculates your average score per nine holes
'defines variables
Dim NameOfCoursex(1 To 500) As String
Dim ParOfCoursex(1 To 500) As Integer
Dim YourScorex(1 To 500) As Integer
Dim NumberOfHolesx(1 To 500) As Integer
Dim Cartx(1 To 500) As String
Dim Datex(1 To 500) As Date
Dim sum As Integer
Dim Counter As Integer
Dim k As Integer
Dim TotalScore As Integer
Dim TotalHolesPlayed As Integer
Dim AvgScore As Single

'hides the totalrecords form and shows the title form
Private Sub cmdBackToHome_Click()
    FrmTotalRecords.Hide
    FrmTitle.Show
End Sub

'hides the totalrecords form and shows the yourrecords form to enter a new record
Private Sub cmdBackToYourRecords_Click()
    FrmTotalRecords.Hide
    FrmYourRecords.Show
End Sub

Private Sub cmdCalculate_Click()
    'opens yourrecords file and goes through it to calculate your average score per nine holes
    Open App.Path & "\YourRecords.txt" For Input As #6
    'initializes variables
    TotalScore = 0
    TotalHolesPlayed = 0
    For k = 1 To Counter
        'if the number of holes played is 18 then it divides the score by 2 to make it a 9 hole basis
        If NumberOfHolesx(k) = 18 Then
            TotalScore = TotalScore + (YourScorex(k) / 2)
        Else
            TotalScore = TotalScore + YourScorex(k)
        End If
    Next k
    'calculates average score per nine holes from the file
    AvgScore = TotalScore / Counter
    picYourTotalRecords.Print ""
    'prints in the picture box what your average score for nine holes is
    picYourTotalRecords.Print "Your average score for nine holes is "; FormatNumber(AvgScore, 2); "."
    'Closes file
    Close #6
End Sub

'ends the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReadYourRecords_Click()
    'clears the picture box and prints the heading and labels for the table
    picYourTotalRecords.Cls
    picYourTotalRecords.Print Tab(2); "Date"; Tab(19); "Name of Golf Course"; Tab(52); "Par for the Course"; Tab(77); "Your Score"; Tab(92); "Number of Holes Played"; Tab(122); "Golf Cart"
    picYourTotalRecords.Print "_________________________________________________________________________________________________________"
    'opens the yourrecords file and saves is to arrays
    Open App.Path & "\YourRecords.txt" For Input As #6
    Counter = 0
    Do While Not EOF(6)
        Counter = Counter + 1
        Input #6, Datex(Counter), NameOfCoursex(Counter), ParOfCoursex(Counter), YourScorex(Counter), NumberOfHolesx(Counter), Cartx(Counter)
        'prints that contents of your file
        picYourTotalRecords.Print Datex(Counter); Tab(19); NameOfCoursex(Counter); Tab(57); ParOfCoursex(Counter); Tab(82); YourScorex(Counter); Tab(101); NumberOfHolesx(Counter); Tab(124); Cartx(Counter)
    Loop
    'closes your file
    Close #6
End Sub
