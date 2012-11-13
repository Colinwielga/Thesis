VERSION 5.00
Begin VB.Form Form4Game1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's hut"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   16
      Top             =   7440
      Width           =   1095
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
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame FrameMenu 
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   9240
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnMainMenu 
         Caption         =   "Main menu"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
   End
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
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6000
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   7200
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   7200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Text            =   "1"
      Top             =   4440
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4815
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "   Now I have a question for you..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "(You already know Person 1 is sitting there...)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   2415
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   120
      Picture         =   "Form4Game1.frx":0000
      Top             =   120
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   $"Form4Game1.frx":A142
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   9615
   End
End
Attribute VB_Name = "Form4Game1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnDone_Click()
    Dim situation1, situation2, situation3, situation4, situation5 As String
    situation1 = Text1.Text
    situation2 = Text2.Text
    situation3 = Text3.Text
    situation4 = Text4.Text
    situation5 = Text5.Text
    

    If situation1 = "" Or (situation2 = "" Or situation3 = "" Or situation4 = "" Or situation5 = "") Then
    MsgBox "Are you kidding? Please fill out all the boxes with numbers."
    ElseIf situation1 = 1 And (situation2 = 3 And situation3 = 4 And situation4 = 5 And situation5 = 2) Then
    MsgBox "Congratulation! Your answer is right!"
    btnNext.Enabled = True
    btnDone.Enabled = False
    Else
    MsgBox "Sorry, please try again."
    End If
End Sub

Private Sub btnExit_Click()
    FormExit.Show
End Sub

Private Sub btnLoad_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #4
        Input #4, stage
        If stage = 4 Then
        Form4Game1.Show
        MsgBox "Loading complete."
        ElseIf stage = 5 Then
        Form4Game1.Hide
        Form5Game2.Show
        MsgBox "Loading complete."
        ElseIf stage = 8 Then
        Form4Game1.Hide
        FinalGame.Show
        MsgBox "Loading complete."
        End If
    
    Close #4
    
End Sub

Private Sub btnMainMenu_Click()
    gotoMain4.Show
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
    Form4Game1.Hide
    Form5Game2.Show
End Sub

Private Sub btnSave_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #1
    Input #1, stage
    Close #1
    
    If stage = 0 Then
    Open App.Path & "\Saves\Save.txt" For Output As #2
    stage = 4
    Write #2, stage
    MsgBox "Game saved."
    Else: FormSaves4.Show
    End If
    
    Close #2
End Sub

