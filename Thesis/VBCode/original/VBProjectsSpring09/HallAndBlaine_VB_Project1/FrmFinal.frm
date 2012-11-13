VERSION 5.00
Begin VB.Form FrmFinal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   7650
   ClientTop       =   4350
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9900
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H000000FF&
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   3960
      ScaleHeight     =   6015
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdLoadArray 
      BackColor       =   &H0000FF00&
      Caption         =   "Did your family make the top 10 high scores?!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "FrmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Family Feud
'frmFinal
'Colin Hall and Andre Blaine
'March 23
'The objective of this form is to show the top ten scores and if the family made it into the top ten
Private Sub cmdEnd_Click()
MsgBox "Thank you for playing Family Feud!", , "See you next time!"
End
End Sub

Private Sub cmdLoadArray_Click()
'This button tells your family if they made it in the top 10
Dim Names(1 To 15) As String, Scores(1 To 15) As Integer, CTR As Integer, X As Integer, Found As Boolean
picResults.Print Tab(7); "Top Ten!"
picResults.Print "_______________________"

Open App.Path & "\scores.txt" For Input As #1     'This loop reads data from a file into two arrays

'Declare the variables
Do While Not EOF(1)     'This loop reads data from a file into two arrays
    CTR = CTR + 1       'Increment the counter
    Input #1, Names(CTR), Scores(CTR)      'Get the next answer and value from the user
Loop
Found = False

'This searches the array and puts the score in the correct place
Do While ((Not Found) And X < CTR)
    X = X + 1
        If Sum > Scores(X) Then
            Found = True
            picResults.Print FamilyName, Sum
        End If
    picResults.Print Names(X), Scores(X)
Loop

Do While X < CTR - 1
    picResults.Print Names(X), Scores(X)
    X = X + 1
Loop
cmdEnd.Visible = True
End Sub

