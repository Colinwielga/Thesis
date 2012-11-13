VERSION 5.00
Begin VB.Form frmRunnerUp 
   BackColor       =   &H8000000D&
   Caption         =   "Runner UP"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   6495
      Left            =   5280
      ScaleHeight     =   6435
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List of Runner Up"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Teams"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   0
      Picture         =   "frmRunnerUp.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   7440
      Width           =   2895
   End
End
Attribute VB_Name = "frmRunnerUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Team(1 To 100) As String
Dim Year(1 To 100) As Single
Dim CTR As Integer


Private Sub cmdLoad_Click()
Open App.Path & "\runnerups.txt" For Input As #3

   
Close #3
End Sub

Private Sub cmdMenu_Click()
frmChampions.Show
frmRunnerUp.Hide
End Sub
'increment ctr each time it goes through the loop to move to the next postion in the array
Private Sub cmdList_Click()
picresults.Cls

    picresults.Print "Team", "Year"
    picresults.Print "*******************************************"
    
    Open App.Path & "\runnerups.txt" For Input As #9
    CTR = 0
    Do While Not EOF(9)

        CTR = CTR + 1
        
        Input #9, Team(CTR), Year(CTR)
        picresults.Print Team(CTR); Tab(30); Year(CTR)
        
    Loop
    
  
    picresults.Print "*****************************************"
   
    Close #9
    
   
    
    End Sub
'This funcion increments the CTR as well as the postion, and is looking through an array to find
'the text provided by the user in the input box
'used the boolean function of found=true or false
'it has 2 arrays
' it is the same formula used in frmGoal
Private Sub cmdSearch_Click()
Dim pos As Integer
    Dim found As Boolean
    Dim teamname As String
    Dim Team(1 To 100) As String
    Dim Years(1 To 100) As Single
    
    Open App.Path & "\runnerups.txt" For Input As #3
    
    picresults.Cls
    
    CTR = 0
    
    Do Until EOF(3)
        CTR = CTR + 1
        Input #3, Team(CTR), Years(CTR)
    Loop
    Close #3
    
    teamname = InputBox("Name a Team that you want to find", "Name")
    
    found = False
    pos = 0
    
    Do While (pos < CTR)
        pos = pos + 1
        If Team(pos) = teamname Then
        picresults.Print teamname; " has been the Runner Up in "; Years(pos)
        found = True
        
        End If
    Loop
    
    If found = True Then
        
    Else
        MsgBox "Sorry this Team has not made it to the  UEFA Champions League Final", , "Sorry"
    End If
    
    Close #7


End Sub


