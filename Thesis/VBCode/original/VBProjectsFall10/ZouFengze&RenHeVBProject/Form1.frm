VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's Hut"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton GodMode 
      Caption         =   "God Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Exitbtn 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton LoadGame 
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton NewGame 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   0
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exitbtn_Click()
    FormExit.Show
End Sub

Private Sub Form_Load()
    WindowsMediaPlayer1.URL = App.Path & "\Sounds\Theme.mp3"
End Sub

Private Sub GodMode_Click()
    'This button was originally considered to show the refences.
    'Since we cited in our project report, we use it as a Test button.
    'For testing any specific form, just motify the code below and use it.
    'That's why I call it "God Mode". I can go to any form directly without triger the story line.

    Form9_2.Show
End Sub

Private Sub LoadGame_Click()
    
    Open App.Path & "\Saves\Save.txt" For Input As #4
        Input #4, stage
        If stage = 4 Then
        Form4Game1.Show
        Form1.Hide
        WindowsMediaPlayer1.Close
        MsgBox "Loading complete."
        
        ElseIf stage = 5 Then
        Form1.Hide
        Form5Game2.Show
        WindowsMediaPlayer1.Close
        MsgBox "Loading complete."
        
        ElseIf stage = 8 Then
        Form1.Hide
        FinalGame.Show
        WindowsMediaPlayer1.Close
        MsgBox "Loading complete."
        
        Else
        MsgBox "Record is empty!"
        End If
    
    Close #4
    
    
End Sub

Private Sub NewGame_Click()
    WindowsMediaPlayer1.Close
    Form1.Hide
    Form2.Show
    
End Sub


