VERSION 5.00
Begin VB.Form FormProfessorSelect 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   12525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16920
   FillColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12525
   ScaleWidth      =   16920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFormSelectProject 
      Caption         =   "Go to Grading Portion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   13
      Top             =   11400
      Width           =   2655
   End
   Begin VB.PictureBox picClass 
      Height          =   4935
      Left            =   6120
      ScaleHeight     =   4875
      ScaleWidth      =   6915
      TabIndex        =   11
      Top             =   6120
      Width           =   6975
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "FormProfessorSelect.frx":0000
      Left            =   5040
      List            =   "FormProfessorSelect.frx":000D
      TabIndex        =   10
      Text            =   "Select Class Period"
      Top             =   1920
      Width           =   6735
   End
   Begin VB.PictureBox picimad 
      Height          =   2175
      Left            =   1800
      Picture         =   "FormProfessorSelect.frx":0067
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   7800
      Width           =   1695
   End
   Begin VB.PictureBox PicMiller 
      Height          =   1935
      Left            =   1800
      Picture         =   "FormProfessorSelect.frx":6769
      ScaleHeight     =   1875
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.PictureBox picLynn 
      Height          =   2055
      Left            =   1800
      Picture         =   "FormProfessorSelect.frx":762C
      ScaleHeight     =   1995
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picTeacher 
      Height          =   855
      Left            =   6240
      ScaleHeight     =   795
      ScaleWidth      =   6795
      TabIndex        =   6
      Top             =   4200
      Width           =   6855
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm Teacher and Class Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.OptionButton Opt_ImadRahal 
      Caption         =   "Imad Rahal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.OptionButton Opt_JohnMiller 
      Caption         =   "John Miller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton Opt_LynnZiegler 
      Caption         =   "Lynn Ziegler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label LblClass 
      Caption         =   "Class List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   12
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label lblProfessor 
      Caption         =   "Select Professor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Caption         =   "CS130 Project Grader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "FormProfessorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
Dim Teacher As String, Period As String, LastNameMiller(1 To 40) As String, FirstNameMiller(1 To 40) As String
Dim CtrMiller As Integer, LastNameLynn(1 To 40) As String, FirstNameLynn(1 To 40) As String, CtrLynn As Integer
Dim LastNameRahal(1 To 40) As String, FirstNameRahal(1 To 40) As String, CtrRahal As Integer

'this if/then statement makes sure the user actaully selects a professor at the beginning of the project
'if the user does not select a professor a message box will tell them that
'once the user selects a professor, it program assigns the name of the professor to the variable "Teacher"
If Opt_LynnZiegler.Value = False _
        And Opt_JohnMiller.Value = False _
        And Opt_ImadRahal.Value = False Then
    MsgBox ("You Must Select a Teacher")
    Opt_LynnZiegler.SetFocus
Else
    If Opt_LynnZiegler.Value = True Then
        Teacher = "Lynn Ziegler"
    ElseIf Opt_JohnMiller.Value = True Then
        Teacher = "John Miller"
    ElseIf Opt_ImadRahal.Value = True Then
        Teacher = "Imad Rahal"
    End If
End If
'this section is where the teacher selects a class
'the user selects the time of the class from the drop down menu
'then if the tacher has a class at that time, then their class list will print in the picture box
Period = cmbClass.Text
If Period = "Period II: 9:40 a.m. (Odd)" And Teacher = "John Miller" Then
    Open App.Path & "\MillerClassData.txt" For Input As #1
CtrMiller = 0
Do Until EOF(1)
    CtrMiller = CtrMiller + 1
    Input #1, LastNameMiller(CtrMiller), FirstNameMiller(CtrMiller)
    picClass.Print FirstNameMiller(CtrMiller), LastNameMiller(CtrMiller)
Loop
Close #1
    ElseIf Period = "Period II: 9:40 a.m. (Odd)" And Teacher = "Lynn Ziegler" Then
        MsgBox ("Lynn Ziegler does not have a class at this time on this day.")
    ElseIf Period = "Period II: 9:40 a.m. (Odd)" And Teacher = "Imad Rahal" Then
        MsgBox ("Imad Rahal does not have a class at this time on this day.")
End If

If Period = "Period III: 11:20 a.m. (Odd)" And Teacher = "Lynn Ziegler" Then
    Open App.Path & "\ZieglerClassData.txt" For Input As #2
CtrLynn = 0
Do Until EOF(2)
    CtrLynn = CtrLynn + 1
    Input #2, LastNameLynn(CtrLynn), FirstNameLynn(CtrLynn)
    picClass.Print FirstNameLynn(CtrLynn), LastNameLynn(CtrLynn)
Loop
Close #2
   
    ElseIf Period = "Period III: 11:20 a.m. (Odd)" And Teacher = "John Miller" Then
        MsgBox ("John Miller does not have a class at this time on this day.")
    ElseIf Period = "Period III: 11:20 a.m. (Odd)" And Teacher = "Imad Rahal" Then
        MsgBox ("Imad Rahal does not have a class at this time on this day.")
End If

If Period = "Period V: 2:40 p.m. (Even)" And Teacher = "Imad Rahal" Then
    Open App.Path & "\RahalClassData.txt" For Input As #3
CtrRahal = 0
Do Until EOF(3)
    CtrRahal = CtrRahal + 1
    Input #3, LastNameRahal(CtrRahal), FirstNameRahal(CtrRahal)
    picClass.Print FirstNameRahal(CtrRahal), LastNameRahal(CtrRahal)
Loop
Close #3
    ElseIf Period = "Period V: 2:40 p.m. (Even)" And Teacher = "John Miller" Then
        MsgBox ("John Miller does not have a class at this time on this day.")
    ElseIf Period = "Period V: 2:40 p.m. (Even)" And Teacher = "Lynn Ziegler" Then
        MsgBox ("Lynn Ziegler does not have a class at this time on this day.")
End If


picTeacher.Print "You Selected "; Teacher; " as the Professor and the "; Period; " Class Period"
        
    
    
End Sub



Private Sub cmdFormSelectProject_Click()
FormProfessorSelect.Hide
FormSelectProject.Show
End Sub

Private Sub Form_Load()
'this sets the optionbuttons to false at the beginning of the form load
Opt_LynnZiegler.Value = False
Opt_JohnMiller.Value = False
Opt_ImadRahal.Value = False
End Sub
