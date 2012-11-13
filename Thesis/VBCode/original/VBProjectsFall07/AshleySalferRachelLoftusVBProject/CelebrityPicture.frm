VERSION 5.00
Begin VB.Form frmCelebrityPicture 
   BackColor       =   &H000000FF&
   Caption         =   "Your Match!"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   9600
      Picture         =   "CelebrityPicture.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox picCupid 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   120
      Picture         =   "CelebrityPicture.frx":14B6
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Your matchmaking experience has ended!  Click to quit."
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2655
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   6555
      TabIndex        =   2
      Top             =   2760
      Width           =   6615
   End
   Begin VB.CommandButton cmdPicture 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click Here If You Are Ready For The Revelation Of Your Match!"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.PictureBox picCelebrityPicture 
      BackColor       =   &H000000FF&
      Height          =   5415
      Left            =   3480
      ScaleHeight     =   5355
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   4080
      Width           =   5775
   End
End
Attribute VB_Name = "frmCelebrityPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPicture_Click()
    'Declares variables to be used in subroutine
    Dim CelebrityName(1 To 100) As String, CelebSex(1 To 100) As String, CelebAge(1 To 100) As String, CelebHair(1 To 100) As String, CelebType(1 To 100) As String, CelebAct(1 To 100) As String
    Dim Counter2 As Integer, Pos As Integer, Found As Boolean
    Dim picturename As String, NoMatch As String
    
    'Opens file and reads selected matches into program
    Open App.Path & "\Celebrity.txt" For Input As #1
    Counter2 = 0
    Do Until EOF(1)
        Counter2 = Counter2 + 1
        Input #1, CelebrityName(Counter2), CelebSex(Counter2), CelebAge(Counter2), CelebHair(Counter2), CelebType(Counter2), CelebAct(Counter2)
    Loop
    Close #1

    Found = False
     
    'Matches user answers to selected answers in file
    Do While (Found = False And Pos < Counter2)
        Pos = Pos + 1
        If UserAnswers(1) = CelebSex(Pos) And UserAnswers(2) = CelebAge(Pos) And UserAnswers(3) = CelebHair(Pos) And UserAnswers(4) = CelebType(Pos) And UserAnswers(5) = CelebAct(Pos) Then
            Found = True
        End If
    Loop
    
    picName.Cls
    
    'If match is found, displays name and picture of celebrity match
    'If no match is found, a warning is shown
    If Found = True Then
        picName.Print "Congratulations! Your match is " & CelebrityName(Pos) & "!"
        picturename = Replace(CelebrityName(Pos), " ", "") & ".jpg"
        picCelebrityPicture = LoadPicture(App.Path & "\Pictures\" & picturename)
    Else
        picName.Print "I'm sorry! We do not have a match for you. Please try again."
    End If
    
End Sub
    
    
Private Sub cmdQuit_Click()

    'Ends program
    End
    
End Sub




Private Sub Form_Load()

End Sub
