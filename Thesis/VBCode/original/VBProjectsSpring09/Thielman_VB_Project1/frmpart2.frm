VERSION 5.00
Begin VB.Form frmpart2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   10485
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14880
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdform3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Take Me To Form Three"
      Height          =   1095
      Left            =   8880
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picresultsthree 
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8640
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdbeginpics 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Begin"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin VB.PictureBox picresults 
      Height          =   6375
      Left            =   1200
      ScaleHeight     =   6315
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label lblscore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbltitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Now It Gets Harder"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmpart2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timberwolves basketball
'frmpart2
'nick thielman
'3/15
'on this form the user will see three pictures and will be asked to id these players/coaches
'message boxes will appear after the user answers the question idication if they are
'right or wrong. The user then is also able to continue onto form 3. This button is
'disabled until the user answers the questions, also there is a runningtotal kept for
'the user's answers and how many they get right.

Private Sub cmdbeginpics_Click()
Dim player As String, runningtotal As Integer
runningtotal = 0
'picture pops up
picresults.Picture = LoadPicture(App.Path & "\kevinlove.jpg")
'input box ask for player's name
player = InputBox("Who is this player?", "Player ID")

'display msgbox for answer
'adds point to runningtotal if correct
If player = "kevin love" Then
runningtotal = runningtotal + 1
picresultsthree.Print runningtotal
Else
MsgBox "That is incorrect, it is kevin love.", , "Wrong"
End If

'new picture pops up
picresults.Picture = LoadPicture(App.Path & "\kevinmchale.jpg")
'input box ask for player's name
player = InputBox("Who is this guy?", "Player ID")

'adds point to runningtotal if correct
If player = "kevin mchale" Then
picresultsthree.Cls
runningtotal = runningtotal + 1
picresultsthree.Print runningtotal
MsgBox "That is correct, it is kevin mchale!", , "Right"
ElseIf player <> "kevin mchale" Then
MsgBox "That is incorrect, it is kevin mchale.", , "Wrong"
End If
'If the user gets two in a row a different message appears
If runningtotal = 2 Then
MsgBox "That is right, two in a row!", , "NICE"
End If

'last picture pops up
'asks user for a name
picresults.Picture = LoadPicture(App.Path & "\mikemiller.jpg")
'input box ask for player's name
player = InputBox("Who is this player?", "Player ID")

'adds point to runningtotal if correct
If player = "mike miller" Then
picresultsthree.Cls
runningtotal = runningtotal + 1
picresultsthree.Print runningtotal
MsgBox "That is correct, it is mike miller!", , "Right"
ElseIf player <> "mike miller" Then
MsgBox "That is incorrect, it is mike miller.", , "Wrong"
End If
'If the user gets three in a row right
If runningtotal = 3 Then
MsgBox "Wow 3 in a row!", , "Great Work"
End If

'tells the user how many pictures they got correct
MsgBox "You were able to identify " & runningtotal & ", player(s)"

  
End Sub

Private Sub cmdform3_Click()
'hides form 2
'goes to form 3
frmpart2.Hide
frmpart3.Show
End Sub
