VERSION 5.00
Begin VB.Form frmMario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MarioCatcher"
   ClientHeight    =   6096
   ClientLeft      =   3156
   ClientTop       =   3732
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6096
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicMario 
      Height          =   612
      Left            =   1440
      Picture         =   "frmMariocatcher.frx":0000
      ScaleHeight     =   564
      ScaleWidth      =   564
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Timer tmrMario 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer tmrGlobal 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   3120
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image imgMario1 
      Height          =   465
      Left            =   1920
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu miStart 
         Caption         =   "&Start"
         Shortcut        =   {F2}
      End
      Begin VB.Menu miStop 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu miSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu miOptions 
         Caption         =   "&Options"
         Shortcut        =   {F5}
      End
      Begin VB.Menu miSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu miExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu miAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "frmMario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmMarioCatcher
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to play a mario catching game.  After selecting a level of dificulty, the user
                'starts the game and clicks on the moving mario picture.  Every correct click is a point.  The total is tallied
                'and the user ends the game by selecting stop from the options.  They can also select a different level of
                'difficulty or view the about section that tells them I made the game.  The last option is to exit the game.

Option Explicit

Public Score As Integer
Public Amount As Integer


Public MarioTimeout As Integer
Public GlobTimeout As Integer

Private GlobTimeoutCount As Integer
Private MarioTimeoutCount As Integer

Private Sub Form_Initialize()
    Dim TempGlobTimeout As String
    Dim TempMarioTimeout As String

    imgMario1 = PicMario
   
    If TempGlobTimeout = "" Or TempMarioTimeout = "" Then
        GlobTimeout = 10
        MarioTimeout = 100
    Else
        TempGlobTimeout = CInt(GetSetting(App.Title, "GameFlags", "GlobTimeout"))
        TempMarioTimeout = CInt(GetSetting(App.Title, "GameFlags", "MarioTimeout"))
    End If
End Sub

Private Sub imgMario1_Click()
   Dim Success As Single
   
   Score = Score + 1
   Success = (Score / Amount)
   
   frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(Success)
   tmrMario.Enabled = False
   MarioTimeoutCount = 0
   imgMario1.Visible = False
   
   tmrGlobal.Enabled = True
   
End Sub

Private Sub miAbout_Click()
    MsgBox "This game was created by Bill Macy", , "About MarioCatcher ..."
End Sub

Private Sub miExit_Click()
    frmMain.Show
    frmOptions.Hide
    frmMariocatcher.Hide
End Sub

Private Sub miStart_Click()
    Dim Success As Single
    Score = 0
    Amount = 0
    Success = (1 / 1)
    MarioTimeoutCount = 0
    frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(Success)

    tmrGlobal.Enabled = True
    
    miStop.Enabled = True
    miOptions.Enabled = False
    miStart.Enabled = False
End Sub

Private Sub miOptions_Click()
    frmOptions.Show
End Sub


Private Sub tmrGlobal_Timer()
    If GlobTimeoutCount = GlobTimeout Then
        GlobTimeoutCount = 0
        
        Randomize
        
        With imgMario1
            .Top = (((frmMariocatcher.ScaleHeight - imgMario1.Height) - 0) * Rnd + 0)
            .Left = (((frmMariocatcher.ScaleWidth - imgMario1.Width) - 0) * Rnd + 0)
            .Visible = True
        End With
        Amount = Amount + 1
        tmrMario.Enabled = True
        tmrGlobal.Enabled = False
    Else
        GlobTimeoutCount = GlobTimeoutCount + 1
    End If
End Sub

Private Sub tmrMario_Timer()
    If MarioTimeoutCount = MarioTimeout Then
        Dim Success As Single
        Success = (Score / Amount)
        frmMariocatcher.Caption = "MarioCatcher - Score: " & Score & " Success Rate: " & FormatPercent(Success)

        MarioTimeoutCount = 0
        tmrGlobal.Enabled = True
        imgMario1.Visible = False
        tmrMario.Enabled = False
    Else
        MarioTimeoutCount = MarioTimeoutCount + 1
    End If
End Sub

