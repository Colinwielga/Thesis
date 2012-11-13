VERSION 5.00
Begin VB.Form records 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   1620
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15240
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   5760
      Picture         =   "records.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   4320
      Picture         =   "records.frx":17F9
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2880
      Picture         =   "records.frx":2391
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   3000
      ScaleHeight     =   555
      ScaleWidth      =   6795
      TabIndex        =   7
      Top             =   7560
      Width           =   6855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1575
   End
   Begin VB.PictureBox picResults1 
      Height          =   5895
      Left            =   8040
      ScaleHeight     =   5835
      ScaleWidth      =   8715
      TabIndex        =   5
      Top             =   1560
      Width           =   8775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Season Records"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton bats 
      BackColor       =   &H000000FF&
      Caption         =   "Batting Records"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Pitching Records"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H000000FF&
      Caption         =   "Press Here For a Fun Fact! "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Twins Records"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   80.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   1
      Top             =   -240
      Width           =   15135
   End
End
Attribute VB_Name = "records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dims variables
Option Explicit
 Dim cord(1 To 30) As String, CTR As Integer, bat(1 To 30) As String, year(1 To 30) As String, cord1(1 To 30) As String, CTR1 As Integer, num(1 To 30) As String, person(1 To 30) As String, cord2(1 To 30) As String, CTR2 As Integer, num2(1 To 30) As String, team2(1 To 30) As String, year2(1 To 30) As String, PLR(1 To 30) As String

 'finds txt file opens it and prints it into the picResults
Private Sub bats_Click()
picResults1.Cls
CTR = 0
   
    
    Open App.Path & "\batting.txt" For Input As #1
    
    
    picResults1.Print "Record", Tab(30); "Bat Avg.", "Year"
    picResults1.Print Tab(1); "*******************************************************************************************************"
    picResults1.Print
    Do While Not EOF(1)
       
        CTR = CTR + 1
        
       Input #1, cord(CTR), bat(CTR), year(CTR)
      
        
        picResults1.Print cord(CTR), Tab(30); bat(CTR), year(CTR)
    Loop
    
   
    picResults1.Print "*******************************************************************************************************************************"
    picResults1.Print
    picResults1.Print
    
    Close #1
End Sub
'opens facts text and then randomizes the info in the file
Private Sub cmdFacts_Click()
Dim Facts(1 To 10) As String, CTR As Integer

Open App.Path & "\facts.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Facts(CTR)
Loop
Close #1
picResults.Cls
picResults.Print Facts(CInt(Int((5 * Rnd()) + 1)))


End Sub
 
'finds txt file opens it and prints it into the picResults
Private Sub Command1_Click()
picResults1.Cls

CTR = 0
   
    
    Open App.Path & "\pitching.txt" For Input As #1
    
    
    picResults1.Print "Record", Tab(35); "Record #", "Player", "Year"
    picResults1.Print Tab(1); "*******************************************************************************************************"
    picResults1.Print
    Do While Not EOF(1)
       
        CTR = CTR + 1
        
       Input #1, cord1(CTR), num(CTR), PLR(CTR), year(CTR)
      
        
        picResults1.Print cord1(CTR), Tab(35); num(CTR), , PLR(CTR), year(CTR)
    Loop
    
   
    picResults1.Print "*******************************************************************************************************************************"
    picResults1.Print
    picResults1.Print
    
    Close #1

End Sub

 'finds txt file opens it and prints it into the picResults

Private Sub Command3_Click()
picResults1.Cls

CTR = 0
   
    
    Open App.Path & "\clubrec.txt" For Input As #1
    
    
    picResults1.Print "Record", Tab(30); "Record #", "Team", "Year"
    picResults1.Print Tab(1); "*******************************************************************************************************"
    picResults1.Print
    Do While Not EOF(1)
       
        CTR = CTR + 1
        
       Input #1, cord2(CTR), num2(CTR), team2(CTR), year2(CTR)
      
        
        picResults1.Print cord2(CTR), Tab(30); num2(CTR), team2(CTR), year2(CTR)
    Loop
    
   
    picResults1.Print "*******************************************************************************************************************************"
    picResults1.Print
    picResults1.Print
    Close #1
End Sub
'shows history form and hides record form
Private Sub Command4_Click()
records.Hide
history.Show

End Sub
'clears picResults
Private Sub Command5_Click()
picResults1.Cls
End Sub
