VERSION 5.00
Begin VB.Form frmSchedule 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   2250
   ClientTop       =   2415
   ClientWidth     =   7470
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7470
   Begin VB.CommandButton cmdA 
      Caption         =   "Return to Main Menu"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H000000FF&
      Caption         =   "Display"
      Height          =   975
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.OptionButton opt2007 
      BackColor       =   &H000000FF&
      Caption         =   "2007"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton opt2006 
      BackColor       =   &H000000FF&
      Caption         =   "2006"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton opt2005 
      BackColor       =   &H000000FF&
      Caption         =   "2005"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox picResults 
      Height          =   6855
      Left            =   2640
      ScaleHeight     =   6795
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblSeason 
      BackColor       =   &H000000FF&
      Caption         =   "Pick the Season You Would Like to View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblSchedule 
      BackColor       =   &H000000FF&
      Caption         =   "Schedule"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdA_Click()
 frmSchedule.Hide
 frmHome.Show
End Sub

Private Sub cmdDisplay_Click()
picResults.Cls

'Declares the variables
Dim Game(1 To 100) As String
Dim Result(1 To 100) As String
Dim A(1 To 100) As String
Dim ctr As Integer
Dim counter As Integer

'Determines which schedule the user wishes to load, and loads it
If opt2005.Value = True Then
    Open App.Path & "\2005.txt" For Input As #1
    Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, A(ctr), Game(ctr), Result(ctr)
    Loop
    Close 1
ElseIf opt2006.Value = True Then
    Open App.Path & "\2006.txt" For Input As #1
    Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, A(ctr), Game(ctr), Result(ctr)
    Loop
    Close 1
ElseIf opt2007.Value = True Then
    Open App.Path & "\2007.txt" For Input As #1
    Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, A(ctr), Game(ctr), Result(ctr)
    Loop
    Close 1
End If


'Prints the schedule
Do While counter < ctr
counter = counter + 1
picResults.Print A(counter), Game(counter), Tab(50), Result(counter)
Loop

 
End Sub
