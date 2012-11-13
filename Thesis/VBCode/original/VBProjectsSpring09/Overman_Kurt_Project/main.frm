VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   2220
   ClientTop       =   2400
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10140
   Begin VB.PictureBox Picture3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000D&
      ForeColor       =   &H8000000D&
      Height          =   4935
      Left            =   3120
      Picture         =   "main.frx":0000
      ScaleHeight     =   4935
      ScaleWidth      =   3615
      TabIndex        =   7
      Top             =   1200
      Width           =   3615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   9000
      Picture         =   "main.frx":843C
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      Picture         =   "main.frx":8B8E
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton team1 
      BackColor       =   &H00FF8080&
      Caption         =   "Team"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton end 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton hist 
      BackColor       =   &H008080FF&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton triv 
      BackColor       =   &H008080FF&
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   $"main.frx":92E0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   8
      Top             =   6240
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Minnesota Twins"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   8040
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'click and program ends
Private Sub end_Click()
End
End Sub
'hides main form and shows history form
Private Sub hist_Click(Index As Integer)
history.Show
main.Hide
End Sub

'hides main form and show team form


Private Sub team1_Click()
team.Show
main.Hide
End Sub
' asks user how much they know about twins trivia and then assigns user to a page depending on ther knowledge
Private Sub triv_Click()
Dim Question1 As String

Question1 = InputBox("How well do you know Twins trivia, Alot Or Not alot?")  'input box is where user types in their answer

If Question1 = "alot" Or Question1 = "Alot" Or Question1 = "alright" Or Question1 = "Alright" Or Question1 = "fine" Or Question1 = "Fine" Then
    Med.Show
main.Hide  'if user types in either those two options then this message box will pop up
    Else
    Easy.Show
    main.Hide  ' if user types in anything other than the two options then this message will pop up

    
    End If



End Sub
