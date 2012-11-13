VERSION 5.00
Begin VB.Form Brainerd 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Match Game"
      Height          =   735
      Left            =   480
      TabIndex        =   19
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to MN!"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdFish2 
      Caption         =   "Fish # 2"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
   End
   Begin VB.PictureBox Picture6 
      Height          =   1695
      Left            =   4800
      ScaleHeight     =   1635
      ScaleWidth      =   3195
      TabIndex        =   7
      Top             =   5400
      Width           =   3255
      Begin VB.CommandButton cmdFish6 
         Caption         =   "Fish # 6"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   3255
      End
      Begin VB.PictureBox Picture7 
         Height          =   1455
         Left            =   0
         Picture         =   "Brainerd.frx":0000
         ScaleHeight     =   1395
         ScaleWidth      =   3075
         TabIndex        =   8
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   1695
      Left            =   600
      Picture         =   "Brainerd.frx":C9D2
      ScaleHeight     =   1635
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   5400
      Width           =   3135
      Begin VB.CommandButton cmdFish3 
         Caption         =   "Fish # 3"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   4800
      Picture         =   "Brainerd.frx":1C734
      ScaleHeight     =   1515
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   3240
      Width           =   3255
      Begin VB.CommandButton cmdFish5 
         Caption         =   "Fish # 5"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   600
      Picture         =   "Brainerd.frx":2ECCE
      ScaleHeight     =   1515
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   3240
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   4800
      Picture         =   "Brainerd.frx":3EA30
      ScaleHeight     =   1395
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
      Begin VB.CommandButton cmdFish4 
         Caption         =   "Fish # 4"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   600
      Picture         =   "Brainerd.frx":50FCA
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton cmdFish1 
         Caption         =   "Fish # 1"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.Label lblAsk 
      Caption         =   "Would you like to play a memory game with fish that are found in MN waters? Follow the directions below!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1695
      Left            =   8400
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"Brainerd.frx":5D99C
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   8400
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblDirections 
      Caption         =   "Directions:          "
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblfishing 
      Caption         =   "Where the up-north, the favorite sport is Fishing!"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label lblBrainerd 
      Caption         =   "  Brainerd "
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Brainerd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesoooota
'Form Name: Brainerd
'Author: Danielle Johnson and Tony Blum
'Date Written: March 26th 2008
'Objective: the user find the matching fish (3--walleye, sunfish, and crappie)

Option Explicit
Dim ONESIX, TWOTHREE, FOURFIVE As Boolean 'declare these variables as global
Dim CTR As Integer

Private Sub cmd_FormLoad()
CTR = 0 'intialize the variables
ONESIX = False
TWOTHREE = False
FOURFIVE = False
End Sub

Private Sub cmdback_Click()
Brainerd.Hide 'Brainerd no longer appears, Minnesota form pops up
Minnesota.Show
End Sub

Private Sub cmdFish1_Click()
CTR = CTR + 1 'represents the first press of the button
    cmdFish1.Visible = False 'pressed button disappears, fish underneath command button will show
    cmdFish2.Visible = True
    cmdFish3.Visible = True
    cmdFish4.Visible = True
    cmdFish5.Visible = True
If TWOTHREE = True Then
    cmdFish2.Visible = False 'show the pictures underneath these command buttons
    cmdFish3.Visible = False
End If
If FOURFIVE = True Then
    cmdFish4.Visible = False 'show the pictures underneath these command buttons
    cmdFish5.Visible = False
End If
If CTR = 2 Then 'second press of the button
    CTR = 0 'cant press any more times, reset ctr
    If cmdFish6.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        ONESIX = True 'the command buttons of 1 and 6 are matching fish, so the above message will pop up
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish1.Visible = True 'the command buttons are visible again
        cmdFish6.Visible = True
    End If
End If

End Sub

Private Sub cmdFish2_Click() 'seek command one commentary, except the switch the fish numbers
CTR = CTR + 1
    cmdFish1.Visible = True
    cmdFish2.Visible = False
    cmdFish4.Visible = True
    cmdFish5.Visible = True
    cmdFish6.Visible = True
If ONESIX = True Then
    cmdFish1.Visible = False
    cmdFish6.Visible = False
End If
If FOURFIVE = True Then
    cmdFish4.Visible = False
    cmdFish5.Visible = False
End If
If CTR = 2 Then
    CTR = 0
    If cmdFish3.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        TWOTHREE = True
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish2.Visible = True
        cmdFish3.Visible = True
    End If
End If

End Sub

Private Sub cmdFish3_Click() 'seek command one commentary, except switch fish numbers
CTR = CTR + 1
    cmdFish1.Visible = True
    cmdFish3.Visible = False
    cmdFish4.Visible = True
    cmdFish5.Visible = True
    cmdFish6.Visible = True
If ONESIX = True Then
    cmdFish1.Visible = False
    cmdFish6.Visible = False
End If
If FOURFIVE = True Then
    cmdFish4.Visible = False
    cmdFish5.Visible = False
End If
If CTR = 2 Then
    CTR = 0
    If cmdFish2.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        TWOTHREE = True
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish2.Visible = True
        cmdFish3.Visible = True
    End If
End If

End Sub

Private Sub cmdFish4_Click() 'seek command one commentary, except switch fish numbers
CTR = CTR + 1
    cmdFish1.Visible = True
    cmdFish2.Visible = True
    cmdFish3.Visible = True
    cmdFish4.Visible = False
    cmdFish6.Visible = True
If TWOTHREE = True Then
    cmdFish2.Visible = False
    cmdFish3.Visible = False
End If
If ONESIX = True Then
    cmdFish1.Visible = False
    cmdFish6.Visible = False
End If
    
If CTR = 2 Then
    CTR = 0
    If cmdFish5.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        FOURFIVE = True
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish4.Visible = True
        cmdFish5.Visible = True
    End If
End If

End Sub

Private Sub cmdFish5_Click() 'seek command one commentary, except switch fish numbers
CTR = CTR + 1
cmdFish1.Visible = True
cmdFish2.Visible = True
cmdFish3.Visible = True
cmdFish3.Visible = True
cmdFish5.Visible = False
cmdFish6.Visible = True
If TWOTHREE = True Then
    cmdFish2.Visible = False
    cmdFish3.Visible = False
End If
If ONESIX = True Then
    cmdFish1.Visible = False
    cmdFish6.Visible = False
End If
If CTR = 2 Then
    CTR = 0
    If cmdFish4.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        FOURFIVE = True
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish4.Visible = True
        cmdFish5.Visible = True
    End If
End If

End Sub

Private Sub cmdFish6_Click() 'seek command one commentary, except switch fish numbers.
CTR = CTR + 1
cmdFish2.Visible = True
cmdFish3.Visible = True
cmdFish3.Visible = True
cmdFish4.Visible = True
cmdFish5.Visible = True
cmdFish6.Visible = False
If TWOTHREE = True Then
    cmdFish2.Visible = False
    cmdFish3.Visible = False
End If
If FOURFIVE = True Then
    cmdFish4.Visible = False
    cmdFish5.Visible = False
End If
If CTR = 2 Then
    CTR = 0
    If cmdFish1.Visible = False Then
        MsgBox "You found the matching pair!", , "YaY!"
        ONESIX = True
    Else
        MsgBox "Sorry. Try Again.", , "Nooo!"
        cmdFish1.Visible = True
        cmdFish6.Visible = True
    End If
End If

End Sub

Private Sub cmdReset_Click()
cmdFish1.Visible = True 'reset the command buttons
cmdFish2.Visible = True
cmdFish3.Visible = True
cmdFish4.Visible = True
cmdFish5.Visible = True
cmdFish6.Visible = True
ONESIX = False 'intialize these variables back to false to play again
TWOTHREE = False
FOURFIVE = False

End Sub
