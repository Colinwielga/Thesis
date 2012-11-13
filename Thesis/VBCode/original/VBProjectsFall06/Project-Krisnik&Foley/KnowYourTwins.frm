VERSION 5.00
Begin VB.Form KnowYourTwins 
   BackColor       =   &H00000080&
   Caption         =   "Know Your Twins"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   3615
      Left            =   840
      Picture         =   "KnowYourTwins.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Click For Player Info."
         Height          =   495
         Left            =   120
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   3615
      Left            =   3960
      Picture         =   "KnowYourTwins.frx":151CE
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Click For Player Info."
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   7080
      Picture         =   "KnowYourTwins.frx":17A78
      ScaleHeight     =   3555
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Click For Player Info."
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Homepage"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   5520
      Picture         =   "KnowYourTwins.frx":1ED1A
      ScaleHeight     =   3435
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   3960
      Width           =   4215
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "Click For Player Info."
         Height          =   495
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   3495
      Left            =   1080
      Picture         =   "KnowYourTwins.frx":215D3
      ScaleHeight     =   3435
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   3960
      Width           =   3975
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "Click For Player Info."
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "KnowYourTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: Twins Baseball
' Form Name: Know Your Twins
' Authors: Jake Krisnik & Mike Foley
' Date Written: October 28, 2006
' Form Objective: To provide the user with the ability to click on a picture to reveal information
'                 about that particular player. The information is linked in our module to public
'                 variables which output the same information that can be found on the Twins roster
'                 form. This gives the user the opportunity to look at a picture of the player and
'                 put a face with the name.
Option Explicit
Dim Ctr As Integer

Private Sub cmdReturn_Click()
' This command button allows the user to navigate away from the Know Your Twins form and return to
' the Homepage.
    HomePage.Show
    KnowYourTwins.Hide
End Sub


Private Sub Command1_Click()
' This command accesses the array we used in a previous form and prints out in a message box
' the player information that relates to the photograph. In particular for this button: Brad
' Radke who occupies the 7th row in our file.
    Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
    Loop
    Close #1                                'Tells the program to pick the appropriate player from the created list
                                            'This then will allow the user to match the player with information about the player
    MsgBox PlayerName(7) & "    " & "#" & PlayerNumber(7) & "    " & PlayerHeight(7) & "In." & "   " & PlayerWeight(7) & "Lbs." & "   " & PlayerPosition(7)
End Sub

Private Sub Command2_Click()
' This command accesses the array we used in a previous form and prints out in a message box
' the player information that relates to the photograph. In particular for this button: Justin
' Morneau who occupies the 16th row in our file.
    Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
    Loop
    Close #1                                    'Tells the program to pick the appropriate player from the created list
                                                'This then will allow the user to match the player with information about the player
    MsgBox PlayerName(16) & "    " & "#" & PlayerNumber(16) & "    " & PlayerHeight(16) & "In." & "   " & PlayerWeight(16) & "Lbs." & "   " & PlayerPosition(16)
End Sub

Private Sub Command3_Click()
' This command accesses the array we used in a previous form and prints out in a message box
' the player information that relates to the photograph. In particular for this button: Torii
' Hunter who occupies the 22nd row in our file.
    Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
    Loop
    Close #1
    MsgBox PlayerName(22) & "    " & "#" & PlayerNumber(22) & "    " & PlayerHeight(22) & "In." & "   " & PlayerWeight(22) & "Lbs." & "   " & PlayerPosition(22)
End Sub

Private Sub Command4_Click()
' This command accesses the array we used in a previous form and prints out in a message box
' the player information that relates to the photograph. In particular for this button: Nick
' Punto who occupies the 18th row in our file.
    Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
    Loop
    Close #1                                        'Tells the program to pick the appropriate player from the created list
                                                    'This then will allow the user to match the player with information about the player
    MsgBox PlayerName(18) & "    " & "#" & PlayerNumber(18) & "    " & PlayerHeight(18) & "In." & "   " & PlayerWeight(18) & "Lbs." & "   " & PlayerPosition(18)
End Sub

Private Sub Command5_Click()
' This command accesses the array we used in a previous form and prints out in a message box
' the player information that relates to the photograph. In particular for this button: Joe
' Mauer who occupies the 12th row in our file.
    Open App.Path & "\Team.txt" For Input As #1
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, PlayerNumber(Ctr), PlayerName(Ctr), PlayerHeight(Ctr), PlayerWeight(Ctr), PlayerPosition(Ctr)
    Loop
    Close #1                                        'Tells the program to pick the appropriate player from the created list
                                                    'This then will allow the user to match the player with information about the player
    MsgBox PlayerName(12) & "    " & "#" & PlayerNumber(12) & "    " & PlayerHeight(12) & "In." & "   " & PlayerWeight(12) & "Lbs." & "   " & PlayerPosition(12)
End Sub
