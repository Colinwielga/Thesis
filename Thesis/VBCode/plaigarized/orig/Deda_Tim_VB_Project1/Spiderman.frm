VERSION 5.00
Begin VB.Form Spiderman 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   7815
   ClientLeft      =   4605
   ClientTop       =   4665
   ClientWidth     =   10860
   LinkTopic       =   "Form2"
   ScaleHeight     =   7815
   ScaleWidth      =   10860
   Begin VB.CommandButton Movie3 
      BackColor       =   &H000000FF&
      Caption         =   "Display Cast For Spider-man 3"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   4080
      ScaleHeight     =   4755
      ScaleWidth      =   6075
      TabIndex        =   5
      Top             =   1440
      Width           =   6135
   End
   Begin VB.CommandButton Movie1 
      BackColor       =   &H000000FF&
      Caption         =   "Display Cast For Spider-man"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Movie2 
      BackColor       =   &H000000FF&
      Caption         =   "Display Cast For Spider-man 2"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton SpiderQuiz 
      BackColor       =   &H000000FF&
      Caption         =   "Continue to Spiderman Quiz"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Read 
      BackColor       =   &H000000FF&
      Caption         =   "Load Lists"
      Height          =   735
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Return 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   8520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Spider-Man Cast Lists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   720
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -2520
      Picture         =   "Spiderman.frx":0000
      Top             =   -120
      Width           =   15360
   End
End
Attribute VB_Name = "Spiderman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare form level variables
Dim SpidermanList1Act(1 To 15) As String, SpidermanList2Act(1 To 15) As String, SpidermanList3Act(1 To 15) As String
Dim SpidermanList1Char(1 To 15) As String, SpidermanList2Char(1 To 15) As String, SpidermanList3Char(1 To 15) As String
Dim CTR As Integer, CTR2 As Integer, CTR3 As Integer

Private Sub Movie1_Click()
'clear other picresults
PicResults.Cls
'print the header info
PicResults.Print "Name of Actor/Actress", "Name of Character"
PicResults.Print "***********************************************************"
'Print cast
Dim I As Integer
PicResults.Print
    For I = 1 To 15
        PicResults.Print SpidermanList1Act(I); Tab(30); SpidermanList1Char(I)
    Next I
End Sub

Private Sub Movie2_Click()
'clear other picresults
PicResults.Cls
'print the header info
PicResults.Print "Name of Actor/Actress", "Name of Character"
PicResults.Print "***********************************************************"
'Print cast
Dim J As Integer
PicResults.Print
    For J = 1 To 15
        PicResults.Print SpidermanList2Act(J); Tab(30); SpidermanList2Char(J)
    Next J
End Sub

Private Sub Movie3_Click()
'clear other picresults
PicResults.Cls
'print the header info
PicResults.Print "Name of Actor/Actress", "Name of Character"
PicResults.Print "***********************************************************"
'Print cast
Dim K As Integer
PicResults.Print
    For K = 1 To 15
        PicResults.Print SpidermanList3Act(K); Tab(30); SpidermanList3Char(K)
    Next K
End Sub

Private Sub Read_Click()
'initialize ctr to 1, to be used for position in the array
CTR = 0
   
'Open and read file
Open App.Path & "\Spiderman Movie.txt" For Input As #1

'set names to variables
    Do While Not EOF(1)
        'set counter and move to next number
        CTR = CTR + 1
        Input #1, SpidermanList1Act(CTR), SpidermanList1Char(CTR)
    Loop
Close #1

'initialize ctr to 1, to be used for position in the array
CTR2 = 0
   
'Open and read file
Open App.Path & "\Spiderman Movie 2.txt" For Input As #2

'set names to variables
    Do While Not EOF(2)
        'set counter and move to next number
        CTR2 = CTR2 + 1
        Input #2, SpidermanList2Act(CTR2), SpidermanList2Char(CTR2)
    Loop
Close #2
        
'initialize ctr to 1, to be used for position in the array
CTR3 = 0
   
'Open and read file
Open App.Path & "\Spiderman Movie 3.txt" For Input As #3

'set names to variables
    Do While Not EOF(3)
        'set counter and move to next number
        CTR3 = CTR3 + 1
        Input #3, SpidermanList3Act(CTR3), SpidermanList3Char(CTR3)
    Loop
Close #3
        
'enable other buttons disable read button
Read.Enabled = False
Movie1.Enabled = True
Movie2.Enabled = True
Movie3.Enabled = True
End Sub

Private Sub Return_Click()
MainMenu.Show
Spiderman.Hide
End Sub

Private Sub SpiderQuiz_Click()
Spider_Quiz.Show
Spiderman.Hide
End Sub
