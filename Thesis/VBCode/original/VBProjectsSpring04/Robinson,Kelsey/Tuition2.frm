VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form2"
   ScaleHeight     =   8775
   ScaleWidth      =   11085
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to find a college"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   5655
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back One Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdDistance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to arrange the colleges by the closes to the cities"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   5655
   End
   Begin VB.CommandButton cmdCost 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to arrange colleges by the cheapest to the most expensive "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to the Next Slide"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton cmdAlpha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to arrange colleges alphabetically"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   3840
      Width           =   8175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Deciding on a College
' Form 2 (Tuition2)
' Kelsey Robinson
' March 10th, 2004
' This form reads the data, sorts it in alphabetical order, finds a college that the user inputs
' It also print the colleges in order of the cheapest to the most expensive, and the closest from the cities to the farthest to the cities.


Option Explicit
Dim X As String


Private Sub Form_Load()
PATH = "N:\CS130\handin\Robinson, Kelsey\"
Open PATH & "colleges.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, College(CTR), Tuition(CTR), Distance(CTR)
    'picResults.Print College(CTR), FormatCurrency(Tuition(CTR)), Distance(CTR)
Loop
Close #1
End Sub

Private Sub cmdAlpha_Click()

picResults.Cls
picResults.Print "Name of the College" '; Tab(35); ; ; "Tuition"; Tab(45); "Distance from the Twin Cities"
picResults.Print "************************************" '**************************************************************"
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
    If College(Comp) > College(Comp + 1) Then
        TempCollege = College(Comp)
        College(Comp) = College(Comp + 1)
        College(Comp + 1) = TempCollege
        TempTuition = Tuition(Comp)
        Tuition(Comp) = Tuition(Comp + 1)
        Tuition(Comp + 1) = TempTuition
        TempDistance = Distance(Comp)
        Distance(Comp) = Distance(Comp + 1)
        Distance(Comp + 1) = TempDistance
    End If
    Next Comp
Next Pass
'picResults.Print "blah"
For J = 1 To CTR
    picResults.Print College(J) '; Tab(30); FormatCurrency(Tuition(J)); Tab(45); , Distance(J)
Next J
End Sub

Private Sub cmdFind_Click()
Found = False
position = 0
X = InputBox("Enter the name of a college")
picResults.Cls
Do While ((Not Found)) And (position < CTR)
    position = position + 1
    If X = College(position) Then
    Found = True
    End If
Loop
If Found Then
        picResults.Print College(position), FormatCurrency(Tuition(position)), Distance(position)
    Else
        picResults.Print "I'm sorry, "; X; " is not in the list of colleges"
End If
End Sub


Private Sub cmdCost_Click()
picResults.Cls
picResults.Print "Name of the College"; Tab(35); "  Tuition" ' Tab(45); "Distance from the Twin Cities"
picResults.Print "********************************************************"
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
    If Tuition(Comp) > Tuition(Comp + 1) Then
        TempTuition = Tuition(Comp)
        Tuition(Comp) = Tuition(Comp + 1)
        Tuition(Comp + 1) = TempTuition
        TempCollege = College(Comp)
        College(Comp) = College(Comp + 1)
        College(Comp + 1) = TempCollege
        TempDistance = Distance(Comp)
        Distance(Comp) = Distance(Comp + 1)
        Distance(Comp + 1) = TempDistance
    End If
    Next Comp
Next Pass
For J = 1 To CTR
    picResults.Print College(J); Tab(35); FormatCurrency(Tuition(J))  'Tab(45); , Distance(J)
Next J

End Sub

Private Sub cmdDistance_Click()
picResults.Cls
picResults.Print "Name of the College"; Tab(30); "Distance from the Twin Cities in miles"
picResults.Print "*********************************************************************************"
For Pass = 1 To CTR - 1
    For Comp = 1 To CTR - Pass
    If Distance(Comp) > Distance(Comp + 1) Then
        TempDistance = Distance(Comp)
        Distance(Comp) = Distance(Comp + 1)
        Distance(Comp + 1) = TempDistance
        TempCollege = College(Comp)
        College(Comp) = College(Comp + 1)
        College(Comp + 1) = TempCollege
        TempTuition = Tuition(Comp)
        Tuition(Comp) = Tuition(Comp + 1)
        Tuition(Comp + 1) = TempTuition
    End If
    Next Comp
Next Pass
For J = 1 To CTR
    picResults.Print College(J); Tab(42); , Distance(J)
Next J
End Sub

Private Sub cmdNext_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub cmdBack_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub



