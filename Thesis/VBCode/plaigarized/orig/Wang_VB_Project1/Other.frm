VERSION 5.00
Begin VB.Form Other 
   Caption         =   "About other Zodiacs"
   ClientHeight    =   9390
   ClientLeft      =   5115
   ClientTop       =   1125
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   Picture         =   "other.frx":0000
   ScaleHeight     =   9390
   ScaleWidth      =   10485
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7320
      TabIndex        =   4
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Main"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   3
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton cmdComments 
      Caption         =   "Some comments about this zodiac"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdCool 
      Caption         =   "Give me a cooler picture!"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picshow 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBack_Click()
other.Visible = False
Home.Visible = True
End Sub

Private Sub cmdComments_Click()
Dim I As Integer, comments(1 To 100) As String, Ctr As Integer
picshow.Cls
Open App.Path & "\Comments\" & Names3(remainder + 1) For Input As #1
Ctr = 0
    Do Until EOF(1)
        Ctr = 1 + Ctr
        Input #1, comments(Ctr)
    Loop
Close #1
    For I = 1 To Ctr
        picshow.Print comments(I)
    Next I
End Sub

Private Sub cmdcool_Click()
picshow.Picture = LoadPicture(App.Path & "\images\" & Names2(remainder + 1))
MsgBox "Honestly, this is the coolest picture of this zodiac I have ever seen!", , "What do you think?"
cmdComments.Enabled = True
End Sub

Private Sub cmdQuit_Click()
End
End Sub

