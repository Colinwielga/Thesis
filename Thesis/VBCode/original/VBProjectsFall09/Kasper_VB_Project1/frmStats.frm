VERSION 5.00
Begin VB.Form frmStatsO 
   Caption         =   "StatsO"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Return to stats"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdInputStats 
      Caption         =   "Tell me more"
      Height          =   375
      Left            =   1440
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   1440
      ScaleHeight     =   1275
      ScaleWidth      =   5715
      TabIndex        =   18
      Top             =   3600
      Width           =   5775
   End
   Begin VB.TextBox txtGName 
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Text            =   " "
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblInstr 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Type a Player's Name by entering the number corresponding to the player"
      Height          =   735
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblCJ 
      BackColor       =   &H00FFFFC0&
      Caption         =   "12. Calvin Johnson"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSmith 
      BackColor       =   &H00FFFFC0&
      Caption         =   "11. Kevin Smith"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblStaff 
      BackColor       =   &H00FFFFC0&
      Caption         =   "10. Matthew Stafford"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblBennett 
      BackColor       =   &H000080FF&
      Caption         =   "9. Earl Bennett"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMF 
      BackColor       =   &H000080FF&
      Caption         =   "8. Matt Forte"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblJC 
      BackColor       =   &H000080FF&
      Caption         =   "7. Jay Cutler"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLions 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lions"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblBears 
      BackColor       =   &H000080FF&
      Caption         =   "Bears"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblDD 
      BackColor       =   &H0000C000&
      Caption         =   "6. Donald Driver"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblRG 
      BackColor       =   &H0000C000&
      Caption         =   "5. Ryan Grant"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblAR 
      BackColor       =   &H0000C000&
      Caption         =   "4. Aaron Rodgers"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblPackers 
      BackColor       =   &H0000C000&
      Caption         =   "Packers"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPercy 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3. Percy Harvin"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblVikings 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vikings"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFavre 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2. Brett Favre"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblPlayers 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1. Adrian Peterson"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   5100
      Left            =   0
      Picture         =   "frmStats.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "frmStatsO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'this form prints stats about the player by taking an input from the user.
'displays results based on the input

Private Sub cmdInputStats_Click()
    Dim Pname As String 'declares Pname as a words
    Pname = txtGName.Text 'sets number entered equal to Pname
    
    If Pname = 1 Then 'condition 1
        picResults.Cls 'clears the picture box
        picResults.Print "Vikings Star running Back Adrian Peterson, rushed for 1760 yars in 2008."
        picResults.Print "So far he has rushed for 7 touchdowns with 481 yards."
        picResults.Print "He was rookie of the year in 2007, and is a complete stud."
     ElseIf Pname = 2 Then 'condition 2
        picResults.Cls 'clears the picture box
        picResults.Print "Known as Judas to most Packer fans, Favre is the oldest Quarterback in the league."
        picResults.Print "He leads the NFC North with 1069 passing yards, and will most likely lead"
        picResults.Print "the Vikings to a Superbowl championship."
     ElseIf Pname = 3 Then 'condition 3
        picResults.Cls 'clears the picture box
        picResults.Print "Rookie from the Florida Gators, Percy Harvin leads the Vikings in receiving yards"
        picResults.Print "with 233 yards, and 2 touchdowns."
     ElseIf Pname = 4 Then 'condition 4
        picResults.Cls 'clears the picture box
        picResults.Print "QB for the Packers, Made first NFL start in the 2008 season after Brett Favre retired."
        picResults.Print "Will most likely lead the Packers to another unsuccessful season. He has passed for"
        picResults.Print "1098 yards so far this season."
     ElseIf Pname = 5 Then 'condition 5
        picResults.Cls 'clears the picture box
        picResults.Print "Nothing compared to Adrian Peterson, but he does his job. Leading the Packers with"
        picResults.Print "257 yards rushing so far on the season, with only two touchdowns though."
     ElseIf Pname = 6 Then 'condition 6
        picResults.Cls 'clears the picture box
        picResults.Print "A packer that can't be hated, Driver leads the team in receiving with"
        picResults.Print "288 yards, and two touchdowns.  Will have a very succesful season"
     ElseIf Pname = 7 Then 'condition 7
        picResults.Cls 'clears the picture box
        picResults.Print "Trying to replace Kyle Orton, Cutler is sneaking by. He has thrown for 901"
        picResults.Print "yards for 8 touchdowns. Best of luck this season."
     ElseIf Pname = 8 Then 'condition 8
        picResults.Cls 'clears the picture box
        picResults.Print "having a slow start but managing, Forte has rushed for 271 yards,"
        picResults.Print "and managed to sneak over the goal line once."
     ElseIf Pname = 9 Then 'condition 9
        picResults.Cls 'clears the picture box
        picResults.Print "Big Earl has 200 yards this season, but has yet to cross the Pilan marker."
     ElseIf Pname = 10 Then 'condition 10
        picResults.Cls 'clears the picture box
        picResults.Print "Suffering from an injury, Stafford better look sharp. He has only 894 yards"
        picResults.Print "passing this season and has thrown for three TDs. Better luck next year,"
        picResults.Print "though he did end the 0-19 streak."
     ElseIf Pname = 11 Then 'condition 11
        picResults.Cls 'clears the picture box
        picResults.Print "With a sad offensive line, Smith has managed to pound out 287 yards and 3 TDs."
     ElseIf Pname = 12 Then 'condition 12
        picResults.Cls 'clears the picture box
        picResults.Print "The go to receiver for young Stafford, Johnson has only 1 touchdown."
        picResults.Print "325 yards is something to brag about though, leading the NFC North."
     Else 'condition 13
        MsgBox "whoops, you must enter a number corresponding with the name.", , "error"
     End If
End Sub

Private Sub cmdquit_Click()
    End
End Sub


Private Sub Command1_Click()
    frmStatsO.Hide 'hides form from user
    FrmOD.Show  'shows form for user
End Sub

