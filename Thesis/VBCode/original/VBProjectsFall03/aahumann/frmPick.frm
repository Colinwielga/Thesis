VERSION 5.00
Begin VB.Form frmPick 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNews 
      Caption         =   "News and Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   18
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Click Here to Calculate scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   17
      Top             =   5760
      Width           =   2295
   End
   Begin VB.OptionButton Bulger 
      BackColor       =   &H000000FF&
      Caption         =   "Marc Bulger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
   End
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      ScaleHeight     =   1755
      ScaleWidth      =   6795
      TabIndex        =   13
      Top             =   6120
      Width           =   6855
   End
   Begin VB.PictureBox pbxPic 
      Height          =   975
      Left            =   720
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.OptionButton Gonzalez 
      BackColor       =   &H000000FF&
      Caption         =   "Tony Gonzalez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton Boston 
      BackColor       =   &H000000FF&
      Caption         =   " David Boston"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
   End
   Begin VB.OptionButton Price 
      BackColor       =   &H000000FF&
      Caption         =   "Peerless Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   9
      Top             =   4800
      Width           =   2535
   End
   Begin VB.OptionButton Vick 
      BackColor       =   &H000000FF&
      Caption         =   " Michael Vick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.OptionButton James 
      BackColor       =   &H000000FF&
      Caption         =   "Edgerrian James"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   4680
      Width           =   2895
   End
   Begin VB.OptionButton Alexander 
      BackColor       =   &H000000FF&
      Caption         =   "Shaun Alexander"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.OptionButton Lewis 
      BackColor       =   &H000000FF&
      Caption         =   "Jamal Lewis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton Williams 
      BackColor       =   &H000000FF&
      Caption         =   "Moe Williams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   4080
      Width           =   3015
   End
   Begin VB.OptionButton Moss 
      BackColor       =   &H000000FF&
      Caption         =   "Randy Moss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.OptionButton Green 
      BackColor       =   &H000000FF&
      Caption         =   "Trent Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblwr 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Wide Receivers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblrb 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Running Backs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblqb 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Quarterbacks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Your Current Team"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmPick (frmPick.frm)
'Purpose of Form: allows users to view the current team
            'form gives users season stats and next weeks opponent
            'also shows news and notes for each player including injuries


Option Explicit



Private Sub Alexander_Click()
'clears stats and picture
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Seahawks"
pbxResults.Print "This Week :  @ Cinncinatti"
pbxResults.Print "This Years stats - 497 Rush Yds. 6 TD's"
pbxPic.Picture = LoadPicture(strPath & "alexander.jpg")
End If
'if player is selected, team name, opponent, and this years stats are diplayed
'picture is also displayed in other picturebox

End Sub

Private Sub Boston_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Chargers"
pbxResults.Print "This Week :  Miami"
pbxResults.Print "This Years stats - 302 Rec Yds. 2 TD's"
pbxPic.Picture = LoadPicture(strPath & "boston.jpg")
End If
End Sub

Private Sub Bulger_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Rams"
pbxResults.Print "This Week :  @ Pit"
pbxResults.Print "This Years stats - 1337 Pass Yds. 10 TD's, 7 INT's"
pbxPic.Picture = LoadPicture(strPath & "bulger.jpg")
End If
End Sub



Private Sub cmdClear_Click()
    
    
End Sub

Private Sub cmdNews_Click()
    
    If Vick = True Then
    MsgBox "News and Notes - Vick probably won't be back Nov. 2nd, his new return date is set at Novemeber 9th "
    ElseIf Bulger = True Then
    MsgBox "Threw for 240, and 3 TD's vs. New York"
    ElseIf Green = True Then
    MsgBox "Thigh Injurt, Probable this week vs. Buffalo"
    ElseIf Alexander = True Then
    MsgBox "Alexander rushed for 120, and 2 TD's vs. Chicago last week"
    ElseIf Lewis = True Then
    MsgBox "Bruised Right Shoulder, Probable vs. Denver"
    ElseIf Williams = True Then
    MsgBox "50 yds rushing, and 50 recieving, with a TD last week vs. Denver"
    ElseIf James = True Then
    MsgBox "Still out with a hurt shoulder, expected back next week vs. Oakland"
    ElseIf Moss = True Then
    MsgBox "114 yds. last week vs. Denver"
    ElseIf Gonzalez = True Then
    MsgBox "74 yds, and a TD vs. Oakland last week"
    ElseIf Boston = True Then
    MsgBox "Horrible game last week, 1 catch for 7 yards"
    ElseIf Price = True Then
    MsgBox "Vick is not expected until November 9th, look for Price's numbers to jump then"
    
    End If
'message box pops up for each player, showing viewer any current news on the player

    
End Sub

Private Sub cmdNext_Click()
frmScores.Visible = True
frmPick.Visible = False
'switches to next form

End Sub

Private Sub cmdQuit_Click()
    End
'ends program
End Sub





Private Sub Form_Load()
    strPath = ("N:\CS130\handin\aahumann\pics\")
End Sub

Private Sub Gonzalez_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Kansas City Chiefs"
pbxResults.Print "This Week :  Buffalo"
pbxResults.Print "This Years stats - 346 Rec Yds. 3 TD's"
pbxPic.Picture = LoadPicture(strPath & "gonzalez.jpg")
End If
End Sub

Private Sub Green_Click()
    pbxResults.Cls
    pbxPic.Cls
 If True Then
 pbxResults.Print "Chiefs"
 pbxResults.Print "This Week :  Buffalo"
 pbxResults.Print "This Years stats - 1562 Pass Yds. 9 TD's, 7 INT's"
 pbxPic.Picture = LoadPicture(strPath & "green.jpg")
 End If
End Sub

Private Sub James_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Colts"
pbxResults.Print "This Week :  Buffalo"
pbxResults.Print "This Years stats - 263 Rush Yds. 1 TD"
pbxPic.Picture = LoadPicture(strPath & "james.jpg")
End If
End Sub

Private Sub Lewis_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Ravens"
pbxResults.Print "This Week : Houston"
pbxResults.Print "This Years stats - 843 Rush Yds. 5 TD's"
pbxPic.Picture = LoadPicture(strPath & "lewis.jpg")
End If
End Sub

Private Sub Moss_Click()

    
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Vikings"
pbxResults.Print "This Week :  Giants"
pbxResults.Print "This Years stats - 666 Rec Yds. 6 TD's"
pbxPic.Picture = LoadPicture(strPath & "moss.jpg")
End If
End Sub

Private Sub Price_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Falcons"
pbxResults.Print "This Week :  BYE"
pbxResults.Print "This Years stats - 347 Rec. Yds. 1 TD"
pbxPic.Picture = LoadPicture(strPath & "price.jpg")
End If
End Sub

Private Sub Vick_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Falcons"
pbxResults.Print "This Week :  BYE"
pbxResults.Print "This Years stats - INJURED"
pbxPic.Picture = LoadPicture(strPath & "vick.jpg")
End If
End Sub

Private Sub Williams_Click()
    pbxResults.Cls
    pbxPic.Cls
If True Then
pbxResults.Print "Vikings"
pbxResults.Print "This Week :  Giants"
pbxResults.Print "This Years stats - 439 Rush Yds. 3 TD's"
pbxPic.Picture = LoadPicture(strPath & "williams.jpg")
End If
End Sub
