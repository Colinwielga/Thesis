VERSION 5.00
Begin VB.Form frmGames2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Games 2"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show Answers!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmddone 
      BackColor       =   &H8000000E&
      Caption         =   "I have put the list in order!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   23
      Top             =   5760
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   2055
      Left            =   5160
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   22
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txt10 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txt9 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txt8 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txt7 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txt6 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txt5 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txt4 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txt3 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox list1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "Games2.frx":0000
      Left            =   360
      List            =   "Games2.frx":0022
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblnames 
      BackStyle       =   0  'Transparent
      Caption         =   "by Lisa Hammer and Kate Bancks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label lbl10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lbl9 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " 9."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lbl8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " 8."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbl7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "7."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  6."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  5."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lbllevel2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Put the List In Alphabetical Order!!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   6705
      Left            =   0
      Picture         =   "Games2.frx":007D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmGames2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ListArray(1 To 10) As String
Dim PlaceArray(1 To 10) As Integer

Private Sub cmddone_Click()                     'the purpose of frmGames2 is to have the user alphabetize a list of words that have been read into an array and sorted into a correct list of answers.
    Dim alphaC As Integer                       'this is level 2 of the game.
                                                'this form keeps of track of how many alphabetized words are in correct order and adds that total to the running total score.
    Dim pos As Integer                          'this button compares input from user by use of text box to a sorted array and increments the score
    Dim placetemp As Integer
    Dim Listtemp As String
    Dim Pass As Integer
    Dim E, F, G, H, I, J, K, L, M, N As String
    E = txt1.Text
    F = txt2.Text
    G = txt3.Text
    H = txt4.Text
    I = txt5.Text
    J = txt6.Text
    K = txt7.Text
    L = txt8.Text
    M = txt9.Text
    N = txt10.Text

    pos = 0
    Open App.Path & "\game2.txt" For Input As #2
    Do Until EOF(2)
        pos = pos + 1
        Input #2, ListArray(pos), PlaceArray(pos)
    Loop
    Close #2

    pos = 0
    For Pass = 1 To (10 - 1)
        For pos = 1 To (10 - Pass)
            If PlaceArray(pos) > PlaceArray(pos + 1) Then
                placetemp = PlaceArray(pos)
                PlaceArray(pos) = PlaceArray(pos + 1)
                PlaceArray(pos + 1) = placetemp
                Listtemp = ListArray(pos)
                ListArray(pos) = ListArray(pos + 1)
                ListArray(pos + 1) = Listtemp
            End If
        Next pos
    Next Pass
        If E = ListArray(1) Then
            alphaC = alphaC + 1
        End If
        If F = ListArray(2) Then
            alphaC = alphaC + 1
        End If
        If G = ListArray(3) Then
            alphaC = alphaC + 1
        End If
        If H = ListArray(4) Then
            alphaC = alphaC + 1
        End If
        If I = ListArray(5) Then
            alphaC = alphaC + 1
        End If
        If J = ListArray(6) Then
            alphaC = alphaC + 1
        End If
        If K = ListArray(7) Then
            alphaC = alphaC + 1
        End If
        If L = ListArray(8) Then
            alphaC = alphaC + 1
        End If
        If M = ListArray(9) Then
            alphaC = alphaC + 1
        End If
        If N = ListArray(10) Then
            alphaC = alphaC + 1
        End If
    MsgBox "Words in Order:" & alphaC, , "Correct"
    C = C + alphaC
    cmdshow.Visible = True
    cmddone.Visible = False



End Sub

Private Sub cmdExit_Click()                                 'the user is allowed to quit the program
    End
End Sub

Private Sub cmdShow_Click()
                                                            'this button shows the sorted list in a picture box and switches forms
    Dim pos As Integer
    For pos = 1 To 10
        picresults.Print PlaceArray(pos), ListArray(pos)
    Next pos
    MsgBox "You've made it to Level 3!", , "WHOOHOO!"
    frmGames2.Hide
    frmseven.Show
    
End Sub

Private Sub Form_Load()                                     'this button reads a file into two parallel arrays
    Open App.Path & "\game2.txt" For Input As #2
    Dim I As Integer
        Do Until EOF(2)
            I = I + 1
            Input #2, ListArray(I), PlaceArray(I)
        Loop
    Close #2
End Sub

