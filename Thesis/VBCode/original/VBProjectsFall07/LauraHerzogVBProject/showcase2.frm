VERSION 5.00
Begin VB.Form showcase2 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00FF80FF&
      Caption         =   "Move on!!!"
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdpass 
      BackColor       =   &H00C00000&
      Caption         =   "Pass!"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdbid 
      BackColor       =   &H00C0C000&
      Caption         =   "Bid!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      Picture         =   "showcase2.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A Trip to Italy!!!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "showcase2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdbid_Click()
'Makes the pass button disabled so that the user can't both bid and pass on a turn
cmdpass.Enabled = False
Dim Bid As Single
Bid = InputBox("Please enter a bid")

If Bid > 25000 Then
    If Bid < 45000 Then
        MsgBox (WholeName) & ("Congratulations, you have won"), , "Winner"
        Runningtotal = Runningtotal + 45000
    Else
        MsgBox (WholeName) & ("You have gone over the price of the items!!"), , "Loser"
        Runningtotal = Runningtotal
    End If
    Else
    MsgBox (WholeName) & (" Your opponent has placed a closer bid that is closer on the items!"), , "Loser"
    Runningtotal = Runningtotal
End If
cmdnext.Visible = True
    


End Sub

Private Sub cmdnext_Click()
Winnings.Show
showcase2.Hide
End Sub

Private Sub cmdpass_Click()
cmdbid.Enabled = False
showcase1.Show
End Sub
