VERSION 5.00
Begin VB.Form frmWinner 
   BackColor       =   &H8000000D&
   Caption         =   "Champions"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form3"
   ScaleHeight     =   7575
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmadrid 
      Caption         =   "The Eternal Champion"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdChampion 
      Caption         =   "Last Champion"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to Main Menu"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Team"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List of Champions"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   6615
      Left            =   4680
      ScaleHeight     =   6555
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmWinner.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   -120
      Width           =   3255
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Team(1 To 100) As String
Dim Year(1 To 100) As Single
Dim CTR As Single

Private Sub cmdChampion_Click()
frmBarca.Show
frmWinner.Hide
End Sub

Private Sub cmdList_Click()

picresults.Cls

picresults.Print "Team", "Year"
    picresults.Print "*******************************************"
    
    Open App.Path & "\Champions.txt" For Input As #14
    CTR = 0
    Do While Not EOF(14)
'increment ctr each time it goes through  the loop to move to the next postion in the array
        CTR = CTR + 1
        
        Input #14, Team(CTR), Year(CTR)
        picresults.Print Team(CTR); Tab(30); Year(CTR)
        
    Loop
    
  
    picresults.Print "*****************************************"
   
    Close #14
    
    End Sub


Private Sub cmdmadrid_Click()
frmMadrid.Show
frmWinner.Hide
End Sub

Private Sub cmdMain_Click()
frmChampions.Show
frmWinner.Hide
End Sub
'This funcion increments the CTR as well as the postion, and is looking through an array to find
'the text provided by the user in the input box
'used the boolean function of found=true or false
'it has 2 arrays
' it is the same formula used in frmGoal
Private Sub cmdSearch_Click()
    
    Dim Pos As Integer
    Dim found As Boolean
    Dim teamname As String
    Dim Team(1 To 100) As String
    Dim Years(1 To 100) As Single
    
    Open App.Path & "\Champions.txt" For Input As #7
    
    picresults.Cls
    
    CTR = 0
    
    Do Until EOF(7)
        CTR = CTR + 1
        Input #7, Team(CTR), Years(CTR)
    Loop
    Close #1
    
    teamname = InputBox("Name a Team that you want to find", "Name")
    
    found = False
    Pos = 0
    
    Do While (Pos < CTR)
        Pos = Pos + 1
        If Team(Pos) = teamname Then
        picresults.Print teamname; " has been Champion in "; Years(Pos)
        found = True
        
        End If
    Loop
    
    If found = True Then
        
    Else
        MsgBox "Sorry this Team has not won the UEFA Champions League", , "Sorry"
    End If
    
    Close #7


End Sub
