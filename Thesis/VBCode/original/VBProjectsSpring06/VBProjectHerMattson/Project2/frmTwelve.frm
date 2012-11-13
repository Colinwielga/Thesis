VERSION 5.00
Begin VB.Form frmHerMattson12 
   BackColor       =   &H000080FF&
   Caption         =   "Search "
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   360
      ScaleHeight     =   5715
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search By Score"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdGender 
      BackColor       =   &H0080FF80&
      Caption         =   "Search by Gender"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H008080FF&
      Caption         =   "Search By Age"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search by Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmHerMattson12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AmazingQuiz
'frmHerMattson12
'Ee Her and Jennifer Mattson
'Written on 3/21/06
'This menu allows user to input key words or numbers and will search for other users who have same information.

Private Sub cmdAge_Click()
    Dim SearchAge As Integer
    Dim J As Single
    Dim Pos As Integer
    Dim Found As Boolean
    picSearch.Cls
    picSearch.Print "First Name", "Last Name", "Age", "Gender", "Score"
    SearchAge = InputBox("Search By Age", "Enter Age")
    For Pos = 1 To Size
        J = InStr(LCase(userage(Pos)), LCase(SearchAge))
        If J <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
        picSearch.Cls
        picSearch.Print "No Ages Found"
    End If
End Sub

Private Sub cmdGender_Click()
    Dim SearchGender As String
    Dim K As String
    Dim Found As Boolean
    Dim Pos As Integer
    picSearch.Cls
    picSearch.Print "First Name", "Last Name", "Age", "Gender", "Score"
    SearchGender = InputBox("Please enter Male (M) or Female (F)", "Search by Gender")
    For Pos = 1 To Size
        K = InStr(LCase(usergender(Pos)), LCase(SearchGender))
        If K <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
    MsgBox "Please enter male or female", , "Incorrect Entry Error"
    SearchGender = InputBox("M or F", "Search by Gender")
        For Pos = 1 To Size
        K = InStr(LCase(usergender(Pos)), LCase(SearchGender))
        If K <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdScore_Click()
    Dim SearchScore As Integer
    Dim E As Single
    Dim Pos As Integer
    Dim Found As Boolean
    picSearch.Cls
    picSearch.Print "First Name", "Last Name", "Age", "Gender", "Score"
    SearchScore = InputBox("Enter a Score", "Score Search")
    For Pos = 1 To Size
        E = InStr(LCase(userscore(Pos)), LCase(SearchScore))
        If E <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
        MsgBox "Please enter a score from 0 to 8", , "Incorrect Entry Error"
        SearchScore = InputBox("Enter a Score", "Score Search")
            For Pos = 1 To Size
            E = InStr(LCase(userscore(Pos)), LCase(SearchScore))
                If E <> 0 Then
                    picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
                    Found = True
                End If
            Next Pos
    End If
End Sub

Private Sub cmdMainMenu_Click()
frmHerMattson12.Hide
frmHerMattson1.Show
End Sub
Private Sub Form_Load()
    Dim Pos As Integer
    Pos = 0
    Open App.Path & "\personal Information.txt" For Input As #2
    Do Until EOF(2)
        Pos = Pos + 1
        Input #2, userlastname(Pos), userfirstname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
    Loop
    Size = Pos
    Close #2
End Sub

Private Sub cmdName_Click()
    Dim SearchName As String
    Dim X As Single
    Dim Y As Single
    Dim Pos As Integer
    Dim Found As Boolean
    picSearch.Print "First Name", "Last Name", "Age", "Gender", "Score"
    SearchName = InputBox("Search for a name", "Name Search")
    For Pos = 1 To Size
        X = InStr(LCase(userlastname(Pos)), LCase(SearchName))
        If X <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    For Pos = 1 To Size
        Y = InStr(LCase(userfirstname(Pos)), LCase(SearchName))
        If Y <> 0 Then
            picSearch.Print userfirstname(Pos), userlastname(Pos), userage(Pos), usergender(Pos), userscore(Pos)
            Found = True
        End If
    Next Pos
    If Found = False Then
        picSearch.Print "No Match Found"
    End If
End Sub

Private Sub picSearch_load()

End Sub

