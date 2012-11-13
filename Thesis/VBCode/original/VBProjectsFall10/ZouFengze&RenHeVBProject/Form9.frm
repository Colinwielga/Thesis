VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   Caption         =   "Witch's hut"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form4"
   ScaleHeight     =   7830
   ScaleWidth      =   10560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton btnDone 
      Caption         =   "Enter the answer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton btnHint 
      Caption         =   "Hint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   120
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
      BackColor       =   &H00000000&
      Caption         =   "I will let you pass..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   """Szkkb Szooldvvm""."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "   When you trying to open the door, it is protected by a spell. And it says, if you can figure out what is "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   6930
      Left            =   0
      Picture         =   "Form9.frx":0000
      Top             =   1320
      Width           =   6750
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDone_Click()
    Dim answer As String
    answer = InputBox("Please enter your answer.", "Answer")
    
    If answer = "Happy Halloween" Then
    WindowsMediaPlayer1.URL = App.Path & "\Sounds\Win.mp3"
    MsgBox "Congratulation! Finally, you get out of the hut."
    Form9.Hide
    Form9_2.Show
    
    Else
    MsgBox "Try again."
    End If
End Sub

Private Sub btnHint_Click()
    MsgBox "If you can figure out the sentenses below, you will get answer. 'When Banshee rh crying, hlnvlmv rh going to die. Yzmhsvv'h power is infinite...'"
End Sub

