VERSION 5.00
Begin VB.Form frmMakeaTeam 
   Caption         =   "Make a Team"
   ClientHeight    =   5850
   ClientLeft      =   4215
   ClientTop       =   2505
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Tw Cen MT Condensed Extra Bold"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   4320
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoadFile 
      Caption         =   "Click Here To Load The File"
      Height          =   1095
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtPlayer5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6960
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlayer1 
      BackColor       =   &H000000FF&
      Caption         =   "Enter Your Starting Five and then click me!"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      MaskColor       =   &H000000C0&
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtPlayer1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblrank2 
      Caption         =   "    Be sure the five players entered are on the Timberwolves roster!!!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5040
      Width           =   8895
   End
   Begin VB.Label lblrank 
      Caption         =   "Enter your starting five and I'll tell you what i think of the lineup!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   -120
      Picture         =   "frmMakeaTeam.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmMakeaTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'I declare these under option explicit so that i can ascces them through the entire form
Dim Player(1 To 16) As String
Dim Rank(1 To 16) As Integer
'I need a counter variable
Dim CTR As Integer
'Sum must be assecible throughout the form
Dim sum As Single

'This Loads the outside file
Private Sub cmdLoadFile_Click()
'I must acces my outside file that ranks players
Open App.Path & "\rankplayers.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Player(CTR), Rank(CTR)
Loop
Close #1
End Sub

Private Sub cmdPlayer1_Click()
sum = 0

'In the previous command i loaded my file now i can use that loaded data to manipulate and show how i want.  In this case the user will input his or her starting five and i will tell them if its a good team
'I will use a match and stop function to retrieve the "rank" of players
'therefore i will need a variable which is the input in the text box, a boolean variable to signify its matched, and a counter variable
Dim Name1 As String
Dim Found1 As Boolean
Dim Pos1 As Integer
Name1 = txtPlayer1
Found1 = False
Do While Found1 = False And Pos1 < CTR
    Pos1 = Pos1 + 1
    If Name1 = Player(Pos1) Then
        Found1 = True
        'the sum variable records the rank of a player, the sum will add together the 5 players rank and tell you if the team is good
        sum = sum + Rank(Pos1)
    End If
Loop
'I will use a match and stop function to retrieve the "rank" of players
'therefore i will need a variable which is the input in the text box, a boolean variable to signify its matched, and a counter variable
Dim Name2 As String
Dim Found2 As Boolean
Dim Pos2 As Integer
Name2 = txtPlayer2
Found2 = False
Do While Found2 = False And Pos2 < CTR
    Pos2 = Pos2 + 1
    If Name2 = Player(Pos2) Then
        Found2 = True
        sum = sum + Rank(Pos2)
    End If
Loop
'I will use a match and stop function to retrieve the "rank" of players
'therefore i will need a variable which is the input in the text box, a boolean variable to signify its matched, and a counter variable
Dim Name3 As String
Dim Found3 As Boolean
Dim Pos3 As Integer
Name3 = txtPlayer3
Found3 = False
Do While Found3 = False And Pos3 < CTR
    Pos3 = Pos3 + 1
    If Name3 = Player(Pos3) Then
        Found3 = True
        sum = sum + Rank(Pos3)
    End If
Loop
'I will use a match and stop function to retrieve the "rank" of players
'therefore i will need a variable which is the input in the text box, a boolean variable to signify its matched, and a counter variable
Dim Name4 As String
Dim Found4 As Boolean
Dim Pos4 As Integer
Name4 = txtPlayer4
Found4 = False
Do While Found4 = False And Pos4 < CTR
    Pos4 = Pos4 + 1
    If Name4 = Player(Pos4) Then
        Found4 = True
        sum = sum + Rank(Pos4)
    End If
Loop
'I will use a match and stop function to retrieve the "rank" of players
'therefore i will need a variable which is the input in the text box, a boolean variable to signify its matched, and a counter variable
Dim Name5 As String
Dim Found5 As Boolean
Dim Pos5 As Integer
Name5 = txtPlayer5
Found5 = False
Do While Found5 = False And Pos5 < CTR
    Pos5 = Pos5 + 1
    If Name5 = Player(Pos5) Then
        Found5 = True
        sum = sum + Rank(Pos5)
    End If
Loop
'All of these different ranks are added together within the sum
'Now i will use a case select and if the sum falls within certain ranges I will show my thoughts on the team

Select Case sum
    Case Is = 27
        MsgBox "This is the best possible lineup you can have!"
    Case Is = 6
        MsgBox "This lineup is terrible, you need to work on your coaching skills!"
    Case 6 To 10
        MsgBox "This is a pretty bad lineup, don't expect to win too many."
    Case 10 To 15
        MsgBox "This lineup is mediocre, maybe they could win a game."
    Case 15 To 20
        MsgBox "This lineup is acceptable, nobody woyuld laugh at it."
    Case 20 To 26
        MsgBox "This lineup is ballin, expect wins and lots of them!"
    Case Else
        MsgBox "Get a full lineup, basketball is a team game. Or Check your spelling."
    End Select
    

        



    
End Sub

Private Sub cmdreturn_Click()
'This allow the user to return to the Timberwolves main page

frmMakeaTeam.Visible = False
frmMainPage.Visible = True

End Sub
