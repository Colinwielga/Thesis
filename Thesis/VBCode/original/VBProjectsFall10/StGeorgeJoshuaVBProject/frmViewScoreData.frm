VERSION 5.00
Begin VB.Form frmViewScoreData 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Student Options - View Score"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   14865
   Begin VB.PictureBox picResponse 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   6855
      Left            =   480
      ScaleHeight     =   6795
      ScaleWidth      =   13275
      TabIndex        =   4
      Top             =   2640
      Width           =   13335
   End
   Begin VB.CommandButton CmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Student Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00000080&
      Caption         =   "Show Score Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox picScores 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1155
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   1320
      Width           =   11415
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Records"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   36
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   10095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "frmViewScoreData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    'Displays the student's grade information and a corresponding message
    'Declares useful varaibles
    Dim grade As String
    Dim pos As Integer
    'Stores the student's position in the pos variables
    pos = StudentPosition
    'Prepares picscores
    picScores.Cls
    picScores.Print " Grade"; Tab(15); "Percent"; Tab(30); "Number Correct"; Tab(50); "Number Wrong"; Tab(70); "Number Attempted"
    picScores.Print "****************************************************************************************************************************************************************************"
    'Gets the letter grade designation and loads the corresponding picture
    Select Case (StudentGrade(pos)) * 100
        Case Is >= 90
            grade = "A"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\Excellent!.jpg")
        Case Is >= 80
            grade = "B"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\Good!.jpg")
        Case Is >= 70
            grade = "C"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\KeepGoing.jpg")
        Case Is >= 60
            grade = "D"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\KeepTrying.jpg")
        Case 0.000001 To 59.9999999
            grade = "F"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\SeekHelp.jpg")
        Case 0
            grade = "N/A"
            picResponse.Picture = LoadPicture(App.Path & "\Pictures\NoGradeYet.jpg")
        Case Else
            MsgBox "Check your Code you messed up"
    End Select
    'Displays the user's grade, percent grade, # correct, # wrong, and # attempted
    picScores.Print " " & grade; Tab(15); FormatPercent(StudentGrade(pos)); Tab(30); studentCorrect(pos); Tab(50); studentWrong(pos); Tab(70); StudentAttempted(pos)
    
    
End Sub

Private Sub cmdLogOut_Click()
    'Logsout the user
    frmViewScoreData.Hide
    Call LogOut
End Sub

Private Sub cmdReturn_Click()
    'Returns the user to the student options pane
    frmViewScoreData.Hide
    frmOptionsPage.Show
End Sub

