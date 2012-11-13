VERSION 5.00
Begin VB.Form Batman_Quiz
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   5115
   ClientTop       =   4500
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   10560
   Begin VB.CommandButton Answer
      BackColor       =   &H8000000C&
      Caption         =   "Guess"
      Height          =   975
      Left            =   2520
      MaskColor       =   &H8000000C&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton NextPic
      BackColor       =   &H8000000C&
      Caption         =   "Next Picture"
      Height          =   975
      Left            =   5400
      MaskColor       =   &H8000000C&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.PictureBox PicResults
      BackColor       =   &H80000012&
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   3000
      ScaleHeight     =   3795
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton MainReturn
      BackColor       =   &H8000000C&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton BatmanReturn
      BackColor       =   &H8000000C&
      Caption         =   "Return to Batman"
      Height          =   735
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label2
      BackStyle       =   0  'Transparent
      Caption         =   "Name the Actor/Actress from the Characters!"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   9975
   End
   Begin VB.Image Image1
      Height          =   15360
      Left            =   0
      Picture         =   "Batman_Quiz.frx":0000
      Top             =   -600
      Width           =   19200
   End
   Begin VB.Label Label1
      Caption         =   "Name That Actor/Actress"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Batman_Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTR As Integer, Ans As String, Batnames(1 To 7) As String, Actnames(1 To 7), I As Integer


Private Sub BatmanReturn_Click()
Batman.Show


Batman_Quiz.Hide
End Sub

Private Sub Form_Load()
'Open the file of picture names and put them in an array called names.

Open App.Path & "\Batnames.txt" For Input As #1
CTR = 0
I = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Batnames(CTR), Actnames(CTR)
Loop
Close #1
End Sub



Private Sub MainReturn_Click()
MainMenu.Show


Batman_Quiz.Hide
End Sub
Private Sub Answer_Click()
Dim C As Integer


C = 0
    Ans = InputBox("Type the name of the actor or actress you think is in the box. Please type both first and last name with proper spacing and caps.")
        If Ans = Actnames(I) Then
            MsgBox ("CORRECT!")


            C = C + 1
        Else
            MsgBox ("Sorry, the right answer was " & Actnames(I) & ". Better luck next time " & UserName & "!")
        End If



        If C = 7 Then

            MsgBox ("Congrats " & UserName & ". You got them all right!")
        End If
End Sub

Private Sub NextPic_Click()


I = I + 1
PicResults.Picture = LoadPicture(App.Path & "\Batman Quiz\" & Batnames(I))
    If I = 7 Then



        NextPic.Enabled = False

    End If

End Sub
