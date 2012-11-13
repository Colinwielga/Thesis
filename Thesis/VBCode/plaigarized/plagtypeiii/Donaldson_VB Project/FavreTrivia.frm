VERSION 5.00
Begin VB.Form frmTrivia
   BackColor       =   &H00004000&
   Caption         =   "The Favre Game"
   ClientHeight    =   10695
   ClientLeft      =   2985
   ClientTop       =   2160
   ClientWidth     =   13515
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form2"
   Picture         =   "FavreTrivia.frx":0000
   ScaleHeight     =   10695
   ScaleWidth      =   13515
   Visible         =   0   'False
   Begin VB.CommandButton cmdPlay
      BackColor       =   &H0000C0C0&
      Caption         =   "Start the Trivia!"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.PictureBox picResults
      BeginProperty Font
         Name            =   "Franklin Gothic Book"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   240
      ScaleHeight     =   6795
      ScaleWidth      =   7875
      TabIndex        =   3
      Top             =   720
      Width           =   7935
   End
   Begin VB.CommandButton cmdAnswers
      BackColor       =   &H00C000C0&
      Caption         =   "Answers"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdMenu
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to Main Menu ==>"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label lblTrivia
      BackStyle       =   0  'Transparent
      Caption         =   "Favre Trivia"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Brett Favre Experience
'frmTrivia
'Doug Donaldson
'2/24/10

'this code will have a few multiple choice questions for the user to answer about
'Brett Favre. Each correct answer earns the user one point
Private Sub cmdPlay_Click()
Dim J As Integer, Ans1 As String, Ans2 As String, Ans3 As String, Ans4 As String, Ans5 As String, Ans6 As String, Ans7 As String, Ans8 As String, Ans9 As String, Ans10 As String
Dim Points As Integer
Points = 0

For J = 1 To 10
    Select Case J
    Case 1
        Ans1 = InputBox("What did Favre major in in college? A. Special Education B. History C. Sports Medicine")
        If Ans1 = "A" Or Ans1 = "a" Then
            Points = Points + 1
        End If
    Case 2
        Ans2 = InputBox("What college did Brett Favre attend? A. LSU B. Mississippi State C. Southern Mississippi")
        If Ans2 = "C" Or Ans2 = "c" Then
            Points = Points + 1
        End If
    Case 3
        Ans3 = InputBox("Brett is missing part of an organ. Name it. A. Kidney B. Lung C. Intestine")
        If Ans3 = "C" Or Ans3 = "c" Then
            Points = Points + 1
        End If
    Case 4
        Ans4 = InputBox("What pick in the NFL was Brett? A. Number 1    B. Number 33    C. Number 64")
        If Ans4 = "B" Or Ans4 = "b" Then
            Points = Points + 1
        End If
    Case 5
        Ans5 = InputBox("What is the significance of November 29th? A. It is Brett's birthday B. It is Brett Favre Day in Wisconsin C. It is his son's birthday")
        If Ans5 = "B" Or Ans5 = "b" Then
            Points = Points + 1
        End If
    Case 6
        Ans6 = InputBox("What is Brett Favre's middle name? A.  Lorenzo B.  Michael C.  David")
        If Ans6 = "A" Or Ans6 = "a" Then
            Points = Points + 1
        End If
    Case 7
        Ans7 = InputBox("What Cameron Diaz movie did Brett make an appearance in? A. Charlie's Angels B. There's Something About Mary C. The Proposal")
        If Ans7 = "B" Or Ans7 = "b" Then
            Points = Points + 1
        End If
    Case 8
        Ans8 = InputBox("What team did Brett play for during the 1992 season? A. Atlanta Falcons B. Green Bay Packers C. Buffalo Bills")
        If Ans8 = "A" Or Ans8 = "a" Then
            Points = Points + 1
        End If
    Case 9
        Ans9 = InputBox("What team did Brett play for in the 2009 season? A. Green Bay Packers B. Minnesota Vikings C. New York Jets")
        If Ans9 = "B" Or Ans9 = "b" Then
            Points = Points + 1
        End If
    Case 10
        Ans10 = InputBox("Who caught Brett's first pass? A. Donald Driver B. Antonio Freeman C. Brett Favre")
        If Ans10 = "C" Or Ans10 = "c" Then
            Points = Points + 1
        End If
    End Select


Next
picResults.Print "Total Points out of 10: "; Points

End Sub


Private Sub cmdAnswers_Click()
picResults.Print "Trivia!"
picResults.Print "------------------------------"
picResults.Print "1. What did Brett Favre major in in college?"
picResults.Print "A. Special Education"
picResults.Print "2. What college did Brett Favre attend?"
picResults.Print "C. Southern Mississippi"
picResults.Print "3. Brett is missing part of one of his organs, name it."
picResults.Print "C. Intestine"
picResults.Print "4. What pick in the NFL draft was Brett Favre?"
picResults.Print "B.  Number 33"
picResults.Print "5. What is the significance of November 29th?"
picResults.Print "B. It is Brett Favre day in Wisconsin"
picResults.Print "6. What is Brett Favre's middle name?"
picResults.Print "A. Lorenzo"
picResults.Print "7. What movie did Brett make an appearance in?"
picResults.Print "B. There's Something About Mary"
picResults.Print "8. What team did Brett play for in the 1992 season?"
picResults.Print "A. Atlanta Falcons"
picResults.Print "9. What team did he play for in the 2009 season?"
picResults.Print "B. Minnesota Vikings"
picResults.Print "10.Who caught Brett's first pass?"
picResults.Print "C. He did."
End Sub



Private Sub cmdMenu_Click()
frmTrivia.Hide
frmMain.Show
End Sub


