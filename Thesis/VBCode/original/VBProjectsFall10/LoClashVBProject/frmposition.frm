VERSION 5.00
Begin VB.Form frmposition 
   BackColor       =   &H00FF0000&
   Caption         =   "position"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbowl 
      Caption         =   "Super Bowl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   9
      Top             =   10320
      Width           =   2055
   End
   Begin VB.CommandButton cmdsuperbowlmvps 
      Caption         =   "Super Bowl MVPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   8
      Top             =   9240
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   7
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton cmdspecial 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Special Teams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      TabIndex        =   6
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmddefense 
      Caption         =   "Defense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12360
      TabIndex        =   5
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdtightend 
      Caption         =   "Tight End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10320
      TabIndex        =   4
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdwidereceiver 
      Caption         =   "Wide Receiver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   3
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdrunningback 
      Caption         =   "Running Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      TabIndex        =   2
      Top             =   8160
      Width           =   1695
   End
   Begin VB.PictureBox picgroup 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   4560
      ScaleHeight     =   5355
      ScaleWidth      =   11355
      TabIndex        =   1
      Top             =   2640
      Width           =   11415
   End
   Begin VB.CommandButton cmdquarterback 
      Caption         =   "Quarterback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label lblmvp 
      BackColor       =   &H00FF0000&
      Caption         =   "Super Bowl MVPs by Position"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4680
      TabIndex        =   10
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbowl_Click()
'switches from form position to form superbowl
frmsuperbowl.Show
frmposition.Hide
End Sub

Private Sub cmddefense_Click()
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim year(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer
Dim pos As String

picgroup.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), player(ctr), position(ctr), team(ctr)
Loop

picgroup.Print "Player    ", , "Team   ", , " Year "
picgroup.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------"
picgroup.Print ; player(7); , ; team(7); , ; year(7)
picgroup.Print ; player(12); , , ; team(12); , ; year(12)
picgroup.Print ; player(14); , ; team(14); , ; year(14)

Close
End Sub

Private Sub cmdquarterback_Click()
'set variables
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim year(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer
Dim pos As String

'clears picture bos
picgroup.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), player(ctr), position(ctr), team(ctr)
Loop

'print
picgroup.Print "Player    ", , "Team   ", , " Year "
picgroup.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------"
picgroup.Print ; player(1); , , ; team(1); , ; year(1)
picgroup.Print ; player(3); , , ; team(3); , ; year(3)
picgroup.Print ; player(4); , , ; team(4); , ; year(4)
picgroup.Print ; player(6); , , ; team(6); , ; year(6)
picgroup.Print ; player(10); , , ; team(10); , ; year(10)
picgroup.Print ; player(11); , ; team(11); , , ; year(11)
picgroup.Print ; player(13); , , ; team(13); , ; year(13)
picgroup.Print ; player(15); , , ; team(15); , ; year(15)
picgroup.Print ; player(18); , ; team(18); , ; year(18)
picgroup.Print ; player(19); , , ; team(19); , ; year(19)
picgroup.Print ; player(21); , , ; team(21); , ; year(21)

Close
End Sub

Private Sub cmdquit_Click()
'quits program
End
End Sub

Private Sub cmdrunningback_Click()
'sets variables
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim year(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer
Dim pos As String

'clears picture box
picgroup.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), player(ctr), position(ctr), team(ctr)
Loop

'prints
picgroup.Print "Player    ", , "Team   ", , " Year "
picgroup.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------"
picgroup.Print ; player(2); , ; team(2); , ; year(2)
picgroup.Print ; player(5); , , ; team(5); , ; year(5)
picgroup.Print ; player(9); , , ; team(9); , ; year(9)

'closes data file so that it can be re-read
Close
End Sub

Private Sub cmdspecial_Click()
'set variables
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim year(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer
Dim pos As String

'clears picture box
picgroup.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), player(ctr), position(ctr), team(ctr)
Loop

'print
picgroup.Print "Player    ", , "Team   ", , " Year "
picgroup.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------"
picgroup.Print ; player(8); , ; team(8); , ; year(8)

'closes data file to be re-read
Close
End Sub

Private Sub cmdsuperbowlmvps_Click()
'switching from the form postion to the form mvp
frmmvp.Show
frmposition.Hide
End Sub

Private Sub cmdtightend_Click()
'clears picture box
picgroup.Cls
'print
picgroup.Print "There were no Super Bowl MVP's at the tight end position"
End Sub

Private Sub cmdwidereceiver_Click()
'set variables
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim year(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer
Dim pos As String

'clears picture box
picgroup.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, year(ctr), player(ctr), position(ctr), team(ctr)
Loop

'print
picgroup.Print "Player    ", , "Team   ", , " Year "
picgroup.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------"
picgroup.Print ; player(16); , ; team(16); , ; year(16)
picgroup.Print ; player(17); , , ; team(17); , ; year(17)
picgroup.Print ; player(20); , ; team(20); , ; year(20)

'closes data file to be re-read
Close
End Sub

Private Sub Label1_Click()

End Sub
