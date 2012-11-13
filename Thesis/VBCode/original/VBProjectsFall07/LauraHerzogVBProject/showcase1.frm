VERSION 5.00
Begin VB.Form showcase1 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   Picture         =   "showcase1.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H000040C0&
      Caption         =   "Move On!!!"
      BeginProperty Font 
         Name            =   "Giddyup Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdpass 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "@Dotum"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdbid 
      BackColor       =   &H000000C0&
      Caption         =   "Bid"
      BeginProperty Font 
         Name            =   "@Kozuka Mincho Pro L"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   3855
      Left            =   360
      Picture         =   "showcase1.frx":0D6E
      ScaleHeight     =   3795
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   3120
      Width           =   5895
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Left            =   3600
      Picture         =   "showcase1.frx":72E8
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox picone 
      Height          =   1575
      Left            =   480
      Picture         =   "showcase1.frx":7E3E
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "showcase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdbid_Click()
'allows the player to place a bid on the items
Dim Bid As Single
Bid = InputBox("please enter a bid", "What is your bid?")
'disables the pass button
cmdpass.Enabled = False
'determines whether the player is a winner
If Bid < 30000 Then
    If Bid > 20000 Then
      MsgBox (WholeName) & (" You have won"), , "WINNER"
      Runningtotal = Runningtotal + 30000
    Else
      MsgBox (WholeName) & (" You have lost; one of your opponents was closer on their bid"), , "Loser"
      Runningtotal = Runningtotal
    End If
Else
    MsgBox (WholeName) & (" You have gone over"), , "Loser"
    Runningtotal = Runningtotal
End If
cmdnext.Visible = True

End Sub

Private Sub cmdnext_Click()
'This button moves on to continue to the next form
showcase1.Hide
Winnings.Show
End Sub

Private Sub cmdpass_Click()
'This command button moves to the other showcase form, it disables the bid function so a user can't choose both options
showcase2.Show
showcase1.Hide
cmdbid.Enabled = False
End Sub

