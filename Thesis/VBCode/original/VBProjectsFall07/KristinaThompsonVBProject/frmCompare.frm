VERSION 5.00
Begin VB.Form frmCompare 
   BackColor       =   &H00000000&
   Caption         =   "Compare"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00808080&
      Caption         =   "Go Back to Answer Page"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9600
      Width           =   2415
   End
   Begin VB.CommandButton cmdWho 
      BackColor       =   &H00808080&
      Caption         =   "Who Should You Study with for the Next Test?"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdDisplayGrade 
      BackColor       =   &H00808080&
      Caption         =   "Display the Scores of the Last Test/Quiz in Descending Order by Grade"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H00808080&
      Caption         =   "Display Practice Scores in Alphabetical Order"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   10575
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00808080&
      Caption         =   "Import The Class Scores"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00808080&
      Height          =   7815
      Left            =   3480
      ScaleHeight     =   7755
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   120
      Picture         =   "frmCompare.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   240
      Picture         =   "frmCompare.frx":67CE4
      ScaleHeight     =   2715
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   5640
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   4455
      Left            =   10080
      Picture         =   "frmCompare.frx":CF9C8
      ScaleHeight     =   4395
      ScaleWidth      =   3195
      TabIndex        =   10
      Top             =   1320
      Width           =   3255
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   9840
      Picture         =   "frmCompare.frx":1376AC
      ScaleHeight     =   2835
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   6000
      Width           =   3255
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   2640
      Picture         =   "frmCompare.frx":19F390
      ScaleHeight     =   1275
      ScaleWidth      =   10755
      TabIndex        =   12
      Top             =   9360
      Width           =   10815
   End
   Begin VB.Label lblHowYouDoing 
      BackColor       =   &H00808080&
      Caption         =   "How Are You Doing In Relation To Your Classmates?"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   10935
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Counts As Single
Dim CMate(1 To 30) As String
Dim PracScor(1 To 30) As Single
Dim LetGrade(1 To 30) As String
Private Sub cmdDisplay_Click()
Dim Pass As Integer
Dim Pos As Integer
Dim Temp As String
Dim K As Integer
picResults.Cls
'put the classmate names in alphabetical order
    For Pass = 1 To Counts - 1
        For Pos = 1 To Counts - Pass
            If CMate(Pos) > CMate(Pos + 1) Then
                Temp = CMate(Pos)
                CMate(Pos) = CMate(Pos + 1)
                CMate(Pos + 1) = Temp
                
                Temp = PracScor(Pos)
                PracScor(Pos) = PracScor(Pos + 1)
                PracScor(Pos + 1) = Temp
                
                Temp = LetGrade(Pos)
                LetGrade(Pos) = LetGrade(Pos + 1)
                LetGrade(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
'this prints the users name and score so they can compare themselves with their classmates
picResults.Print YourName, TrySum
picResults.Print "  "
picResults.Print "Name "; "  Practice Scores"
picResults.Print "*****************************"
'print the name and practice scores of the classmates in alphabetical order
For K = 1 To Counts
        picResults.Print CMate(K), PracScor(K)
Next K
End Sub
Private Sub cmdDisplayGrade_Click()
Dim Pass As Integer
Dim Pos As Integer
Dim Temp As String
Dim K As Integer
picResults.Cls
'print users name and score so they can compare with the classmates
picResults.Print YourName, LastScore
picResults.Print "              "
'order the names of the classmates by lowest grade to highest grade
    For Pass = 1 To Counts - 1
        For Pos = 1 To Counts - Pass
            If LetGrade(Pos) < LetGrade(Pos + 1) Then
                Temp = LetGrade(Pos)
                LetGrade(Pos) = LetGrade(Pos + 1)
                LetGrade(Pos + 1) = Temp
                
                Temp = PracScor(Pos)
                PracScor(Pos) = PracScor(Pos + 1)
                PracScor(Pos + 1) = Temp
                
                Temp = CMate(Pos)
                CMate(Pos) = CMate(Pos + 1)
                CMate(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
picResults.Print "Name "; " Grade"
picResults.Print "*****************************"
'print the classmates name and letter grade
For K = 1 To Counts
        picResults.Print CMate(K), LetGrade(K)
Next K
End Sub
Private Sub cmdGoBack_Click()
'this allows you to go back to the previous form
frmAnswer.Show
frmCompare.Hide
frmStudyGuide.Hide
frmWelcome.Hide
End Sub
Private Sub cmdImport_Click()
Open App.Path & "\Classmates.txt" For Input As #3
    Do While Not EOF(3)
        Counts = Counts + 1
        Input #3, CMate(Counts), PracScor(Counts), LetGrade(Counts)
        'picResults.Print CMate(Counts), PracScor(Counts), LetGrade(Counts)
    Loop
Close #3
'this makes sure the user presses the correct button first and doesn't press it again
cmdDisplayGrade.Visible = True
cmdDisplay.Visible = True
cmdWho.Visible = True
cmdImport.Visible = False
End Sub
Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdWho_Click()
'this button tells the user who to study with based on the user score
    MsgBox "Study with anyone whose grade was higher than " & LastScore & " so they can help you improve your previous score"
End Sub
