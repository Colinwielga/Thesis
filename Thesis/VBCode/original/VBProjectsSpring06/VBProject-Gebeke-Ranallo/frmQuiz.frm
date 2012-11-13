VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00FFFFFF&
   Caption         =   "What's Your Style?"
   ClientHeight    =   9495
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "frmQuiz.frx":0000
   ScaleHeight     =   9495
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSortCategory 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort by Style"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSortAlphabetically 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort Alphabetically"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00FF8080&
      Caption         =   "Display Previous Users"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      FillColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7320
      ScaleHeight     =   1995
      ScaleWidth      =   6795
      TabIndex        =   6
      Top             =   2760
      Width           =   6855
   End
   Begin VB.TextBox txtNameSearch 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00FF8080&
      Caption         =   "Take Quiz NOW!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblNameLookup 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Name"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblQuizTakers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Look up the results of previous quiz takers: "
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblNames 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenna Gebeke ~ Katie Ranallo"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   3
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblQuiz 
      BackStyle       =   0  'Transparent
      Caption         =   "What's Your Style?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Quiz
'Form Objective: This form allows the user to navigate to the quiz questions or to return to the startup page.
'Form Obj. cont.: This form also allows the user to search the file of previous quiz takers and allows them to sort the file.
Option Explicit
Dim size As Integer
Dim class(1 To 50) As String
Dim pos As Integer

'This command button displays the file of previous quiz takers.
Private Sub cmdDisplay_Click()
'opening the file of previous quiz takers
Open App.Path & "\QuizTakers.txt" For Input As #1
    picResults.Cls
    pos = 0
    Do Until EOF(1)
        pos = pos + 1
        Input #1, UserNames(pos), class(pos)
        picResults.Print UserNames(pos), class(pos)
    Loop
    Close #1
    size = pos
End Sub

Private Sub cmdQuiz_Click()
'This command button allows the user to navigate to the quiz questions.
    frmQuiz.Hide
    frmTakeQuiz1.Show
End Sub


Private Sub cmdReturn_Click()
'This command button allows the user to return to the startup page.
    frmStart.Show
    frmQuiz.Hide
End Sub
'This command button searches for a user input within the file of previous quiz takers.
Private Sub cmdSearch_Click()
Dim NameSearch As String
Dim Found As Boolean
Dim size As Integer
Dim pos As Integer
Dim count As Integer
count = 0
pos = 0
Dim catogory As String
    NameSearch = txtNameSearch.Text
    
    'opening the file
    Open App.Path & "\QuizTakers.txt" For Input As #1
    picResults.Cls
    Do Until EOF(1)
        pos = pos + 1
        Input #1, UserNames(pos), class(pos)
    Loop
    Close #1
    size = pos
    pos = 0
    
    'searching the file for a match
    Do While (Found = False) And (pos < size)
        count = count + 1
        If NameSearch = UserNames(count) Then
            Found = True
        End If
        pos = pos + 1
    Loop
    'printing the results of the search
    If Found = True Then
        picResults.Print UserNames(count), class(pos)
    Else
        picResults.Print "Match not Found"
    End If
        

End Sub
'The command button sorts the previous users alphabetically.
Private Sub cmdSortAlphabetically_Click()
Dim pass As Integer
Dim Temp1, Temp2 As String
Dim n As Single
For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If UserNames(pos) > UserNames(pos + 1) Then
                Temp1 = UserNames(pos)
                UserNames(pos) = UserNames(pos + 1)
                UserNames(pos + 1) = Temp1
                Temp2 = class(pos)
                class(pos) = class(pos + 1)
                class(pos + 1) = Temp2
            End If
        Next pos
    Next pass
    picResults.Cls
    For n = 1 To size
    picResults.Print UserNames(n), class(n)
    Next n
    
End Sub
'The command button sorts the styles of previous users together.
Private Sub cmdSortCategory_Click()
Dim pass As Integer
Dim Temp1, Temp2 As String
Dim n As Single
For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If class(pos) > class(pos + 1) Then
                Temp1 = UserNames(pos)
                UserNames(pos) = UserNames(pos + 1)
                UserNames(pos + 1) = Temp1
                Temp2 = class(pos)
                class(pos) = class(pos + 1)
                class(pos + 1) = Temp2
            End If
        Next pos
    Next pass
    picResults.Cls
    For n = 1 To size
    picResults.Print UserNames(n), class(n)
    Next n
End Sub

Private Sub Form_Load()
    frmQuiz.Caption = "Welcome " & userName & "  - What's Your Style?"
    
End Sub


