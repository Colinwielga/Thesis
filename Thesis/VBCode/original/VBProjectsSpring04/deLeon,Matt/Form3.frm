VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Alphabetize"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdg3 
      Caption         =   "Go to Form 3"
      Height          =   855
      Left            =   7440
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdform2 
      Caption         =   "Go to Form 2"
      Height          =   855
      Left            =   7440
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox picresults 
      Height          =   5295
      Left            =   2520
      ScaleHeight     =   5235
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton Cmdread 
      Caption         =   "Read Roster"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3300
      Left            =   7320
      Picture         =   "Form3.frx":0000
      Top             =   720
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:TeamStats stats.vbg
'Author Matt de Leon
'Form Name: Form 1 (Form3.frm)
'Written March 15 2004
'Purpose of form: to display a roster of one team, showing Players names, numbers and positions.
'Also to allow the user to sort the players in alphabetical order.
'Throughout the forms the user will be allowed to find stats on certain players and input stats on players and compare them to their averages.
'user will be allowed to sort alphabetically and by scoring averages of players.

Option Explicit
Dim Player(1 To 30) As String, Path As String
Dim CTR As Integer

Private Sub cmdAlpha_Click()
Dim NewName As String
Dim PASS As Integer, COMP As Integer, J As Integer
CTR = 12

'Arrange Alphabetically
For PASS = 1 To CTR - 1
For COMP = 1 To CTR - PASS
If Player(COMP) > Player(COMP + 1) Then
        
'switch names
NewName = Player(COMP)
Player(COMP) = Player(COMP + 1)
Player(COMP + 1) = NewName
            
End If
Next COMP
Next PASS

picresults.Print
picresults.Print
picresults.Print "____________________Alphabetical Order_______________________________"

'Show Players in Alphabetical order
For J = 1 To CTR
    picresults.Print Player(J)
Next J
End Sub

Private Sub cmdclear_Click()
picresults.Cls
End Sub

Private Sub cmdform2_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub cmdg3_Click()
Form3.Show
Form1.Hide
Form2.Hide

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Cmdread_Click()
 'initialize ctr to zero, to be used for position in the array
 CTR = 0
   
    'Open file to be used
    Open Path & "roster.txt" For Input As #1
    picresults.Print "ROSTER"
    picresults.Print "Player", "No.", "Position"
    picresults.Print "----------------------------------------------------------------"
    Do While Not EOF(1)
        'increase ctr each time through the loop to move to the next spot
    
        CTR = CTR + 1
        
        'Put data form file into the array and print data
        Input #1, Player(CTR)
        picresults.Print Player(CTR); Tab(1)
           
    Loop
  
   Close #1
    
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\deLeon,Matt\"
End Sub
