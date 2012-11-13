VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form5Game2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's hut"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10710
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnMenuOpen 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame FrameMenu 
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   9240
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton btnExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnMainMenu 
         Caption         =   "Main menu"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnMenuClose 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton btnDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton btnHint 
      Caption         =   "Hint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   6360
      Width           =   4935
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Text            =   "21"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   11
         Text            =   "9"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Text            =   "51"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Text            =   "11"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Text            =   "6"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Text            =   "8"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "  This way..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   4920
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "   The lock looks like this:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   5940
      Left            =   5880
      Picture         =   "Form5Game2.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4710
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   $"Form5Game2.frx":22F8F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   120
      Picture         =   "Form5Game2.frx":23052
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "Form5Game2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDone_Click()
    Dim situation1, situation2, situation3, situation4, situation5, situation6, situation7, situation8 As String
    situation1 = Text1.Text
    situation2 = Text2.Text
    situation3 = Text3.Text
    situation4 = Text4.Text
    situation5 = Text5.Text
    situation6 = Text6.Text
    situation7 = Text7.Text
    situation8 = Text8.Text
    
    If situation1 = "" Or (situation2 = "" Or situation3 = "" Or situation4 = "" Or situation5 = "" Or situation6 = "" Or situation7 = "" Or situation8 = "") Then
    MsgBox "Please fill out all the boxes with numbers!"
    ElseIf situation1 = 1 And (situation2 = 1 And situation3 = 2 And situation4 = 3 And situation5 = 5 And situation6 = 8 And situation7 = 13 And situation8 = 21) Then
    MsgBox "The door unlocked!"
    btnNext.Enabled = True
    btnDone.Enabled = False
    btnHint.Enabled = False
    Else
    MsgBox "Nothing happened."
    End If
End Sub


Private Sub btnHint_Click()
    MsgBox "On the backside of the lock, there's a sentense: 1, 3, and 21 are my best friends forever."
End Sub

Private Sub btnMenuClose_Click()
    FrameMenu.Visible = False
    btnSave.Visible = False
    btnLoad.Visible = False
    btnMainMenu.Visible = False
    btnExit.Visible = False
    btnMenuClose.Visible = False
    btnMenuOpen.Visible = True

End Sub

Private Sub btnMenuOpen_Click()

    FrameMenu.Visible = True
    btnSave.Visible = True
    btnLoad.Visible = True
    btnMainMenu.Visible = True
    btnExit.Visible = True
    btnMenuOpen.Visible = False
    btnMenuClose.Visible = True
End Sub

Private Sub btnNext_Click()
    Form5Game2.Hide
    Form6.Show
    
End Sub

Private Sub Form_Load()
    WindowsMediaPlayer1.URL = App.Path & "\Sounds\Annie_this_way.wav"
End Sub
Private Sub btnExit_Click()
    FormExit.Show
End Sub

Private Sub btnLoad_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #4
        Input #4, stage
        If stage = 4 Then
        Form4Game1.Show
        Form5Game2.Hide
        MsgBox "Loading complete."
        ElseIf stage = 5 Then
        Form5Game2.Show
        MsgBox "Loading complete."
        ElseIf stage = 8 Then
        Form5Game2.Hide
        FinalGame.Show
        MsgBox "Loading complete."
        End If
    
    Close #4
    
End Sub

Private Sub btnMainMenu_Click()
    gotoMain5.Show
End Sub
Private Sub btnSave_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #1
    Input #1, stage
    Close #1
    
    If stage = 0 Then
    Open App.Path & "\Saves\Save.txt" For Output As #2
    stage = 5
    Write #2, stage
    MsgBox "Game saved."
    Else: FormSaves5.Show
    End If
    
    Close #2
End Sub

