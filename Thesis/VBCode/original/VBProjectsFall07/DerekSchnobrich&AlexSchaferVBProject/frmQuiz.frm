VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   2505
   ClientTop       =   4125
   ClientWidth     =   7800
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdHome 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   4320
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txt5 
      Height          =   735
      Left            =   5040
      TabIndex        =   10
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txt4 
      Height          =   735
      Left            =   5040
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txt3 
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txt2 
      Height          =   735
      Left            =   5040
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txt1 
      Height          =   735
      Left            =   5040
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lbl5 
      Caption         =   "John Gagliardi coached what other sport at St. John's?"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label lbl4 
      Caption         =   "St. John's Leads Division 3 in what non-team related category anually?"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label lbl3 
      Caption         =   "John Gagliardi is famous for winning with ________. (Fill in the blank)"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label lbl2 
      Caption         =   "What year did John Gagliardi take over as head coach?"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lbl1 
      Caption         =   "When was the  Johnnies last National Championship?"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Returns user to main menu
Private Sub cmdHome_Click()
frmQuiz.Hide
frmHome.Show
End Sub

Private Sub cmdSubmit_Click()
cmdSubmit.Enabled = False 'Keeps user from making multiple entries
'Defines Variables
Dim Answer1 As Integer
Dim Answer2 As Integer
Dim Answer3 As String
Dim Answer4 As String
Dim Answer5 As String
Dim Correct As Integer

'Reads in answers
Answer1 = txt1.Text
Answer2 = txt2.Text
Answer3 = txt3.Text
Answer4 = txt4.Text
Answer4 = txt4.Text

'Checks Answers
If Answer1 = 2003 Then
    Correct = Correct + 1
End If

If Answer2 = 1956 Then
    Correct = Correct + 1
End If

If Answer3 = No Then
    Correct = Correct + 1
End If

If Answer4 = Attendance Then
    Correct = Correct + 1
End If

If Answer5 = Hockey Then
    Correct = Correct + 1
End If

'Lets user know their score
MsgBox ("You got " & Correct & " correct")

End Sub
