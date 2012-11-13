VERSION 5.00
Begin VB.Form Frmsearch 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Frmsearch.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtActor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   5400
      Width           =   4095
   End
   Begin VB.CommandButton CmdActor 
      Caption         =   "Search by Actor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtChar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox Txtmovie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton CmdChar 
      Caption         =   "Search By Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Cmdmenu 
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Cmdmovie 
      Caption         =   "Search By Movie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   4680
      Width           =   2415
   End
   Begin VB.PictureBox PicResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2520
      ScaleHeight     =   4275
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter names with capital letters. (e.g. ""Cinderella"")"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   6480
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter names with capital letters. (e.g. ""Will Smith"")"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   5760
      Width           =   5055
   End
End
Attribute VB_Name = "Frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim actors(1 To 100) As String
Dim movies(1 To 100) As String
Dim Characters(1 To 100) As String
Dim Ctr As Integer
Dim N As String

'Movie Trivia
'FrmBeauty
'Amber Olson, Emily Borka, Shannon O'Neill
'11-1-08
'The purpose of this form is to allow the user to search for their favorite star or movie and print them alphabetically.

Private Sub CmdActor_Click()
'This is to search array by actor name.
Dim B As Integer
Dim C As Integer
Dim Mov(1 To 100) As String
Dim Act(1 To 100) As String
Dim Char(1 To 100) As String
Dim matchesCtr As Integer
Dim actorvalue As Boolean

'Initialization of variables.

matchesCtr = 0
PicResults.Cls
N = txtActor.Text
actorvalue = False

'This is a loop to search the array for the match to the movie and stores values to temp arrays for printing purposes.


For B = 1 To Ctr
    If actors(B) = N Then
        matchesCtr = matchesCtr + 1
        Mov(matchesCtr) = movies(B)
        Act(matchesCtr) = actors(B)
        Char(matchesCtr) = Characters(B)
        actorvalue = True
    End If
Next B


If actorvalue = True Then
    For C = 1 To matchesCtr
        PicResults.Print Mov(C); Tab(30); Char(C)
    Next C
Else
    MsgBox "No Results Found", , "Epic Fail"
End If

End Sub

Private Sub CmdChar_Click()
'This is to search array by character name.
Dim B As Integer
Dim C As Integer
Dim Mov(1 To 100) As String
Dim Act(1 To 100) As String
Dim Char(1 To 100) As String
Dim matchesCtr As Integer
Dim charvalue As Boolean

'Initialization of variables.

matchesCtr = 0
PicResults.Cls
N = txtChar.Text
charvalue = False

'This is a loop to search the array for the match to the movie and stores values to temp arrays for printing purposes.


For B = 1 To Ctr
    If Characters(B) = N Then
        matchesCtr = matchesCtr + 1
        Mov(matchesCtr) = movies(B)
        Act(matchesCtr) = actors(B)
        Char(matchesCtr) = Characters(B)
        charvalue = True
    End If
Next B

If charvalue = True Then
    For C = 1 To matchesCtr
        PicResults.Print Mov(C); Tab(30); Act(C)
    Next C
Else
    MsgBox "No Results Found", , "Epic Fail"
End If


End Sub

Private Sub Cmdmenu_Click()
'This is to get back to the main menu and get all of the other forms to hide.
FrmForm1.Show
FrmForm2.Hide
FrmForm3.Hide
FrmForm4.Hide
Frmsearch.Hide
End Sub

Private Sub Cmdmovie_Click()
'This is to search array by movie title.
Dim B As Integer
Dim C As Integer
Dim Mov(1 To 100) As String
Dim Act(1 To 100) As String
Dim Char(1 To 100) As String
Dim matchesCtr As Integer
Dim movvalue As Boolean

'Initialization of variables.

matchesCtr = 0
PicResults.Cls
N = Txtmovie.Text
movvalue = False

'This is a loop to search the array for the match to the movie and stores values to temp arrays for printing purposes.

For B = 1 To Ctr
    If movies(B) = N Then
        matchesCtr = matchesCtr + 1
        Mov(matchesCtr) = movies(B)
        Act(matchesCtr) = actors(B)
        Char(matchesCtr) = Characters(B)
        movvalue = True
    End If
Next B


'This is a bubble sort to display character names and actor names alphabetical by character name.

Dim Pass As Integer
Dim Pos As Integer
Dim Temp As String
Dim Temp2 As String

For Pass = 1 To matchesCtr - 1
    For Pos = 1 To matchesCtr - Pass
        If Char(Pos) > Char(Pos + 1) Then
            Temp = Char(Pos)
            Char(Pos) = Char(Pos + 1)
            Char(Pos + 1) = Temp
            Temp2 = Act(Pos)
            Act(Pos) = Act(Pos + 1)
            Act(Pos + 1) = Temp2
        End If
    Next Pos
Next Pass

'This is where it prints if match is found. If not a message box pops up.

If movvalue = True Then
    For C = 1 To matchesCtr
        PicResults.Print Mov(C); Tab(30); Act(C); Tab(60); Char(C)
    Next C
Else
    MsgBox "No Results Found", , "Epic Fail"
End If
    

End Sub



Private Sub Form_Load()
'To load the array from notepad file when page opens.
Ctr = 0
PicResults.Cls
Open App.Path & "\movieactors.txt" For Input As #1
    Do Until EOF(1)
        Ctr = Ctr + 1
            Input #1, actors(Ctr), Characters(Ctr), movies(Ctr)
    Loop
Close #1

        
End Sub


