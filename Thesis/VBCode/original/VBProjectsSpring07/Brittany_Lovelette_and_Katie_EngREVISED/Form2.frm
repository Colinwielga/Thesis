VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00400040&
   Caption         =   "Home"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHoroscope 
      BackColor       =   &H008080FF&
      Caption         =   "Your Horoscope"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoCalculators 
      BackColor       =   &H008080FF&
      Caption         =   "Birthday Calculator"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdBirthstone 
      BackColor       =   &H008080FF&
      Caption         =   "Your Birthstone"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      Height          =   1335
      Left            =   1200
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00C0C0FF&
      Height          =   5175
      Left            =   4800
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit 'spell check, checks for errors, Cited: Lecture 11
    
'This button moves the user from the form frmHome to the form frmBirthstone.
'Cited: Lecture 18
Private Sub cmdBirthstone_Click()
    frmBirthstone.Show
    'makes the form frmBirthstone visible to the user
    frmHome.Hide
    'makes the form frmHome invisible to the user.
End Sub

'This button takes the user to the form frmCalculators.
'Cited: Lab 7, Problem 3
Private Sub cmdGoCalculators_Click()
    frmHome.Visible = False
    'makes the form frmHome invisible to the user
    frmCalculators.Visible = True
    'makes the form frmCalculators visible to the user
End Sub

Private Sub cmdHoroscope_Click()
    frmHome.Hide
    'makes the form frmHome invisible to the user
    frmHoroscopes.Show
    'makes the form frmHoroscopes visible to the user
    'This button takes the user to the form frmHoroscopes
    'Cited: Lecture 18
End Sub

'This button ends the program. It also displays a message box with the online sources we used for our project.
'Cited: Lecture 11
Private Sub cmdQuit_Click()
    Dim Temp As String
    Dim Pass As Integer
    Dim References(1 To 10) As String
    'Declares variables
    'Cited: Lecture 11
    Open App.Path & "\references2.txt" For Input As #1
    'opens file
    CTR = 0
    'initializes CTR as zero
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, References(CTR)
    'reads data from file into an array titled References
    Loop
    'Above cited: Imad's VB Programs, "Average Exam Scores"
        
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If References(Pos) > References(Pos + 1) Then
                Temp = References(Pos)
                References(Pos) = References(Pos + 1)
                References(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    'The above is a Bubble Sort to alphabatize the sources listed in the message box
    'Cited: Imad's VB Programs, "Sort Project"
            
    MsgBox "References:" & vbNewLine & References(1) & vbNewLine & References(2) & vbNewLine & References(3) & vbNewLine & References(4) & vbNewLine & References(5) & vbNewLine & References(6) & vbNewLine & References(7) & vbNewLine & References(8), , "References"
    'displays a message box to the user listing the sources before exiting the program.
    'Cited: TA Chris Kerber and Lecture 12
    
    End
    'ends the program
    'Cited: Lecture 11
End Sub


