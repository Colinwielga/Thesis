VERSION 5.00
Begin VB.Form frmMatchingGame 
   BackColor       =   &H00FF0000&
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12"
      Height          =   1335
      Index           =   12
      Left            =   7560
      TabIndex        =   30
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11"
      Height          =   1335
      Index           =   11
      Left            =   7560
      TabIndex        =   29
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10"
      Height          =   1335
      Index           =   10
      Left            =   7560
      TabIndex        =   28
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "9"
      Height          =   1335
      Index           =   9
      Left            =   7560
      TabIndex        =   27
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      Height          =   1335
      Index           =   8
      Left            =   7560
      TabIndex        =   26
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      Height          =   1335
      Index           =   7
      Left            =   7560
      TabIndex        =   25
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      Height          =   1335
      Index           =   6
      Left            =   360
      TabIndex        =   24
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      Height          =   1335
      Index           =   5
      Left            =   360
      TabIndex        =   23
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      Height          =   1335
      Index           =   4
      Left            =   360
      TabIndex        =   22
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      Height          =   1455
      Index           =   3
      Left            =   360
      TabIndex        =   21
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      Height          =   1335
      Index           =   2
      Left            =   360
      TabIndex        =   20
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJU 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Height          =   1335
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox lblResults 
      BackColor       =   &H0080FF80&
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Text            =   "Results"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOtherForm 
      BackColor       =   &H000080FF&
      Caption         =   "Go To Record Book"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdEnterName 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter Your Name"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FFFF&
      Height          =   2535
      Left            =   3720
      ScaleHeight     =   2475
      ScaleWidth      =   2835
      TabIndex        =   13
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1335
      Left            =   3720
      TabIndex        =   12
      Top             =   8280
      Width           =   3135
   End
   Begin VB.PictureBox picTwins2 
      Height          =   1095
      Left            =   7920
      Picture         =   "Matching.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   8280
      Width           =   1935
   End
   Begin VB.PictureBox picColliseum2 
      Height          =   1215
      Left            =   7920
      Picture         =   "Matching.frx":086E
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   10
      Top             =   6600
      Width           =   2055
   End
   Begin VB.PictureBox picColliseum 
      Height          =   1215
      Left            =   7920
      Picture         =   "Matching.frx":182E
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox picMountains2 
      Height          =   1095
      Left            =   7920
      Picture         =   "Matching.frx":27EE
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picTwins 
      Height          =   1095
      Left            =   7920
      Picture         =   "Matching.frx":364A
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox picTower2 
      Height          =   1335
      Left            =   8040
      Picture         =   "Matching.frx":3EB8
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox picPyrimads2 
      Height          =   1215
      Left            =   600
      Picture         =   "Matching.frx":47C6
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   8400
      Width           =   2055
   End
   Begin VB.PictureBox picSJU2 
      Height          =   1335
      Left            =   840
      Picture         =   "Matching.frx":534E
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
   End
   Begin VB.PictureBox picPyrimads 
      Height          =   1335
      Left            =   600
      Picture         =   "Matching.frx":64DF
      ScaleHeight     =   1275
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   5040
      Width           =   2175
   End
   Begin VB.PictureBox picTower 
      Height          =   1455
      Left            =   720
      Picture         =   "Matching.frx":7067
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.PictureBox picMountains 
      Height          =   1215
      Left            =   720
      Picture         =   "Matching.frx":7975
      ScaleHeight     =   1155
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.PictureBox picSJU 
      Height          =   1335
      Left            =   840
      Picture         =   "Matching.frx":87D1
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblTitle2 
      BackColor       =   &H000080FF&
      Caption         =   "Matching Game!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   18
      Top             =   6840
      Width           =   3615
   End
   Begin VB.Label lblTitle1 
      BackColor       =   &H00FF00FF&
      Caption         =   "James and Mike's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   17
      Top             =   5640
      Width           =   3855
   End
End
Attribute VB_Name = "frmMatchingGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Dim Ctr As Integer
Dim UserName As String
Dim First As Boolean
Dim picturenames(1 To 12) As String
Dim FirstChosen As Integer
Dim Count1 As Integer


Private Sub cmdEnterName_Click()
    
    UserName = InputBox("Please Enter Your First Name", "Enter Name")
    
End Sub

Private Sub cmdOtherForm_Click()
    frmMatchingGame.Visible = False
    frmSort.Visible = True
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSJU_Click(Index As Integer)
    
    Dim i As Single, score As Integer
    First = Not First
    score = Ctr
    
    
    If First Then
        cmdSJU(Index).Visible = False
        FirstChosen = Index
        Ctr = Ctr + 1
    ElseIf picturenames(Index) = picturenames(FirstChosen) Then
        cmdSJU(Index).Visible = False
        Count1 = Count1 + 1
        Ctr = Ctr + 1
    Else
        cmdSJU(Index).Visible = False
        
        Sleep 100
        
        cmdSJU(Index).Visible = True
        cmdSJU(FirstChosen).Visible = True
        Ctr = Ctr + 1
    End If
    
    If Count1 = 6 Then
        picresults.Print UserName; "Your Score is "; score
        
        Select Case score
            Case Is <= 35
                MsgBox ("Great Job")
            Case 36 To 46
                MsgBox ("Good Job But There Is Room For Improvement")
            Case 47 To 60
                MsgBox ("You Need Practice")
            Case Else
                MsgBox ("Find A New Hobby")
        End Select
    End If
    
        
End Sub

Private Sub Form_Load()
    Dim RandomNumber As Single, Temp As String, j As Integer
    First = False
    Ctr = 0
    Open App.Path & "\picturenames.txt" For Input As #1
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, picturenames(Ctr)
    Loop
    Close #1
    
    Randomize
   
    
    For j = 1 To 3
        For i = 1 To 12
            RandomNumber = Int(Rnd * 11 + 1)
            Temp = picturenames(i)
            picturenames(i) = picturenames(RandomNumber)
            picturenames(RandomNumber) = Temp
        Next i
    Next j
    
    picSJU.Picture = LoadPicture(App.Path & "\" & picturenames(1))
    picMountains.Picture = LoadPicture(App.Path & "\" & picturenames(2))
    picTower.Picture = LoadPicture(App.Path & "\" & picturenames(3))
    picPyrimads.Picture = LoadPicture(App.Path & "\" & picturenames(4))
    picSJU2.Picture = LoadPicture(App.Path & "\" & picturenames(5))
    picPyrimads2.Picture = LoadPicture(App.Path & "\" & picturenames(6))
    picTower2.Picture = LoadPicture(App.Path & "\" & picturenames(7))
    picTwins.Picture = LoadPicture(App.Path & "\" & picturenames(8))
    picMountains2.Picture = LoadPicture(App.Path & "\" & picturenames(9))
    picColliseum.Picture = LoadPicture(App.Path & "\" & picturenames(10))
    picColliseum2.Picture = LoadPicture(App.Path & "\" & picturenames(11))
    picTwins2.Picture = LoadPicture(App.Path & "\" & picturenames(12))
   
   
  
    
    

End Sub

