VERSION 5.00
Begin VB.Form Instructions 
   BackColor       =   &H0000FF00&
   Caption         =   "Instructions"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox HelpBox 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   6735
   End
   Begin VB.CommandButton Ready 
      BackColor       =   &H00FFFF00&
      Caption         =   "I am ready!  Let's play the Palonzi Piano!"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   7095
   End
   Begin VB.CommandButton Helpp 
      BackColor       =   &H00FFFF00&
      Caption         =   "Help me, please."
      Height          =   615
      Left            =   3720
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Help 
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "5) Closing the program"
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "4) Pressing the keys"
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "3) Getting the keyboard"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "2) Entering a beginning note."
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "1) Starting"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Cap 
      BackColor       =   &H0000FF00&
      Caption         =   "What do you need help with?"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Palonzison Piano
'This is the Instructions Form
'Matthew Peterson and Nicholas Alonzi are the authors of this Form
'This form was written in 2009 in the month of March
'This form is written to allow the user to "ask" questions to the program and specify what the user needs help with
    'instead of overwhelming them with too much potentially not needed information.
    
Private Sub Helpp_Click()
Dim HelpNum As Integer, HelpMe As String
    HelpBox = " "
    HelpNum = Help.Text
    If HelpNum > Ctr Or HelpNum < 1 Then
        MsgBox ("Please enter a number between 1 and 5.")
    Else
        HelpMe = InsNum(HelpNum) & ") " & Instruction(HelpNum)
        HelpBox = HelpMe
    End If
End Sub

Private Sub Ready_Click()
Dim TempNote(1 To 99) As String
    Open App.Path & "\notefiles.txt" For Input As #1
    Ctr = 0
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, TempNote(Ctr)
        Notes(Ctr) = App.Path & TempNote(Ctr)
    Loop
    Close #1
    Instructions.Hide
    Piano.Show
End Sub
