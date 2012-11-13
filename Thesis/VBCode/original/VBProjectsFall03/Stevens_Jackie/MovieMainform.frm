VERSION 5.00
Begin VB.Form MovieMain 
   BackColor       =   &H00800080&
   Caption         =   "Movies: Main"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort Movies"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FF0000&
      Caption         =   "Get Movie Info"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton cmdEndofDay 
      BackColor       =   &H00FF0000&
      Caption         =   "End of Day Totals"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00FF0000&
      Caption         =   "List Screens"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtScreen 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      BackColor       =   &H00FF0000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   2295
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
      Height          =   6255
      Left            =   3120
      ScaleHeight     =   6195
      ScaleWidth      =   7515
      TabIndex        =   8
      Top             =   120
      Width           =   7575
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtSenior 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtChildren 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtGeneral 
      BeginProperty Font 
         Name            =   "ModernBlck"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblScreen 
      BackColor       =   &H00800080&
      Caption         =   "Screen Number:"
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
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblPass 
      BackColor       =   &H00800080&
      Caption         =   "Passes:"
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
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label LblSeniors 
      BackColor       =   &H00800080&
      Caption         =   "Seniors:"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblChildren 
      BackColor       =   &H00800080&
      Caption         =   "Children:"
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblGeneral 
      BackColor       =   &H00800080&
      Caption         =   "General Admission:"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "MovieMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: MovieProject (MoveProject.vbp)
'Form Name: MovieMain (MovieMainform.frm)
'Author: Jackie Stevens
'Date Written: 10/20/03
'Purpose of form: 1. To calculate the cost of movie tickets per order:
                    'user will input screen number, and ticket information
                    'and the program will display the total cost according to
                    'the users information.
                '2. To provide links to other forms in the program that provide
                    'different information the user may need.
'Purpose of project: To create a program that could be utilized in an actual
                    'movie theater.  It will display costs for tickets, end of day
                    'totals, movie information, and easy ways to search for movies.
                    'This program is designed so that it could be used either by a
                    'customer or employee at a movie theater.
                
Option Explicit

Private Sub cmdEndofDay_Click()
    'Go to EndofDay Form
EndOfDay.Show
MovieMain.Hide
End Sub

Private Sub cmdInfo_Click()
    'Go to MovieInfo Form
MovieInfo.Show
MovieMain.Hide

End Sub

Private Sub cmdList_Click()
    'Initialize Counter
CTR = 0
    'Open File
Open Path & "MovieFile.txt" For Input As #1
    'Clear Screen
picResults.Cls
    'Print a list of movies to see what screen they are playing in
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Screen(CTR), Movie(CTR), Rating(CTR)
    picResults.Print Screen(CTR); Tab(7); Movie(CTR)
Loop
Close

End Sub

Private Sub cmdQuit_Click()
    'Quits program
End
End Sub

Private Sub cmdSort_Click()
    'Go to MovieSort Form
MovieSort.Show
MovieMain.Hide
End Sub

Private Sub cmdTotal_Click()
    'Declare all local variables
Dim General As Integer, Children As Integer, Seniors As Integer, Passes As Integer
Dim Total As Single, GeneralTotal As Single, ChildrenTotal As Single, SeniorsTotal As Single
Dim ScreenNumber As String
    'Initialize Counter
CTR = 0
    'Open File
Open Path & "MovieFile.txt" For Input As #1
    'Read into array
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Screen(CTR), Movie(CTR), Rating(CTR)
Loop
Close
    'Initialize all other local variables
ScreenNumber = Val(txtScreen.Text)
General = Val(txtGeneral.Text)
Children = Val(txtChildren.Text)
Seniors = Val(txtSenior.Text)
Passes = Val(txtPass.Text)
GeneralTotal = General * 7
ChildrenTotal = Children * 5
SeniorsTotal = Seniors * 5
Total = 0
    'Begin loop to display message if screen number is invalid and ask for valid number
Do While ScreenNumber < 1 Or ScreenNumber > 18
    MsgBox "Screen Number must be between 1 and 18", , "Invalid Entry"
    ScreenNumber = InputBox("Enter a screen number between 1 and 18", "Screen Number")
Loop
    'Clear Screen
picResults.Cls
    'Find total for individual order and add it into running total
Total = GeneralTotal + ChildrenTotal + SeniorsTotal
EndTotal = EndTotal + Total
    'Prints movie order is for
picResults.Print Screen(ScreenNumber), Movie(ScreenNumber)
   
    'Puts total into total for that movie during the day
If ScreenNumber = 1 Then
        MovieTotal1 = MovieTotal1 + Total
    ElseIf ScreenNumber = 2 Then
        MovieTotal2 = MovieTotal2 + Total
    ElseIf ScreenNumber = 3 Then
        MovieTotal3 = MovieTotal3 + Total
    ElseIf ScreenNumber = 4 Then
        MovieTotal4 = MovieTotal4 + Total
    ElseIf ScreenNumber = 5 Then
        MovieTotal5 = MovieTotal5 + Total
    ElseIf ScreenNumber = 6 Then
        MovieTotal6 = MovieTotal6 + Total
    ElseIf ScreenNumber = 7 Then
        MovieTotal7 = MovieTotal7 + Total
    ElseIf ScreenNumber = 8 Then
        MovieTotal8 = MovieTotal8 + Total
    ElseIf ScreenNumber = 9 Then
        MovieTotal9 = MovieTotal9 + Total
    ElseIf ScreenNumber = 10 Then
        MovieTotal10 = MovieTotal10 + Total
    ElseIf ScreenNumber = 11 Then
        MovieTotal11 = MovieTotal11 + Total
    ElseIf ScreenNumber = 12 Then
        MovieTotal12 = MovieTotal12 + Total
    ElseIf ScreenNumber = 13 Then
        MovieTotal13 = MovieTotal13 + Total
    ElseIf ScreenNumber = 14 Then
        MovieTotal14 = MovieTotal14 + Total
    ElseIf ScreenNumber = 15 Then
        MovieTotal15 = MovieTotal15 + Total
    ElseIf ScreenNumber = 16 Then
        MovieTotal16 = MovieTotal16 + Total
    ElseIf ScreenNumber = 17 Then
        MovieTotal17 = MovieTotal17 + Total
    ElseIf ScreenNumber = 18 Then
        MovieTotal18 = MovieTotal18 + Total
End If
    'Prints summary and total for individual order
picResults.Print "-------------------------------------------------------------"
picResults.Print "General("; General; ")"; Tab(17); "="; Tab(23); FormatCurrency(GeneralTotal)
picResults.Print "Children("; Children; ")"; Tab(17); "="; Tab(23); FormatCurrency(ChildrenTotal)
picResults.Print "Seniors("; Seniors; ")"; Tab(17); "="; Tab(23); FormatCurrency(SeniorsTotal)
picResults.Print "Passes("; Passes; ")"; Tab(17); "="; Tab(23); "$0.00"
picResults.Print "-------------------------------------------------------------"
picResults.Print "Total"; Tab(17); "="; Tab(23); FormatCurrency(Total)

End Sub

