VERSION 5.00
Begin VB.Form frmGoal 
   BackColor       =   &H8000000D&
   Caption         =   "Top Scorers"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   5160
      Picture         =   "frmGoal.frx":0000
      ScaleHeight     =   2655
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   4800
      Width           =   4695
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to main Menu"
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdRaul 
      Caption         =   "Top Scorer"
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdlistGoal 
      Caption         =   "List"
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton CmdPlayer 
      Caption         =   "Player"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1935
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4335
      Left            =   4560
      ScaleHeight     =   4275
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   0
      Picture         =   "frmGoal.frx":B598
      ScaleHeight     =   4875
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmGoal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Player(1 To 100) As String
Dim Goals(1 To 100) As Single
Dim Playername As String
Dim found As Boolean
Dim K As Single

Private Sub Form_Load()
picResults2.Visible = False

End Sub

Private Sub cmdLoad_Click()
Open App.Path & "\topscorer.txt" For Input As #5
  CTR = 0
    
    Do While Not EOF(5)
'increment ctr each time in the loop the loop and then moves it to the next postion in the array
        CTR = CTR + 1
        
        Input #5, Player(CTR), Goals(CTR)
        Loop
        End Sub


Private Sub cmdMenu_Click()
frmChampions.Show
frmGoal.Hide

End Sub

Private Sub cmdlistGoal_Click()

picresults.Cls

picresults.Print "Player", "Goals"
    picresults.Print "*******************************************"
    
    
        For K = 1 To CTR
        picresults.Print Player(K); Tab(30); Goals(K)
        Next K
        

        
    
        
    picresults.Print "*****************************************"
       
End Sub

'This funcion increments the CTR as well as the postion, and is looking through an array to find
'the text provided by the user in the input box
'used the boolean function of found=true or false
'it has 2 arrays
'part of the formula was taken from the VBProject Bancks&Hammer


Private Sub CmdPlayer_Click()
Dim Pos As Integer
    Dim found As Boolean
    Dim Playerinput As String
    Dim Player(1 To 100) As String
    Dim Goals(1 To 100) As Single
    
    Open App.Path & "\topscorer.txt" For Input As #2
    
    CTR = 0
    
    picresults.Cls
    
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, Player(CTR), Goals(CTR)
    Loop
    Close #10
    
    Playerinput = InputBox("Name a Player that you want to find", "Name")
    
    found = False
    Pos = 0
    
    Do While (Pos < CTR)
        Pos = Pos + 1
        If Player(Pos) = Playerinput Then
        picresults.Print Playerinput; " has scored "; Goals(Pos); "in the UEFA Champions league"
        found = True
        
        End If
    Loop
    
    If found = True Then
        
    Else
        MsgBox "Sorry this Player is not a top Scorer UEFA Champions League", , "Sorry"
    End If
    
    Close #2
End Sub

Private Sub cmdRaul_Click()
picResults2.Visible = True
'shows a picture once the button is pressed
End Sub
