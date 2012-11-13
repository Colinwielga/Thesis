VERSION 5.00
Begin VB.Form Details
   Caption         =   "Details of your Chinese Zodiac"
   ClientHeight    =   9060
   ClientLeft      =   4245
   ClientTop       =   1470
   ClientWidth     =   10425
   Picture         =   "Details.frx":0000
   ScaleHeight     =   9060
   ScaleWidth      =   10425
   Begin VB.CommandButton cmdBack
      Caption         =   "Back to Main"
      BeginProperty Font
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
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
      Height          =   1575
      Left            =   7800
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdComments
      Caption         =   "Any comments about my zodiac?"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdcool
      Caption         =   "Give me a cooler picture of my zodiac!"
      BeginProperty Font
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picshow
      AutoSize        =   -1  'True
      BeginProperty Font
         Name            =   "Vivaldi"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project: Chinese Zodiac
' Form name: Details
' Author: Haosen Wang
' Date: Feb. 24, 2010
' Objective: this form is to give some pictures of the zodiac, and some comments about it.
Private Sub cmdBack_Click()
Details.Visible = False
Details.Visible = False
Details.Visible = False
Home.Visible = True
MsgBox "Glad to see you again! Little " & zodiac(remainder + 1), , "Welcome Back!"
End Sub

Private Sub cmdComments_Click()
Dim I As Integer, comments(1 To 100) As String, Ctr As Integer
picshow.Cls
Open App.Path & "\Comments\" & Names3(remainder + 1) For Input As #1  'this allows the program to read different files
Ctr = 0                                                               'according to what the designated zodiac is.
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
MsgBox "Honestly, this is the coolest picture of your zodiac I have ever seen!", , "What do you think?"
cmdComments.Enabled = True              'Activating the other form.

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Dim I As Integer
MsgBox "So, you think this the picture of your Zodiac is silly? Click on the first button and have a cooler one!", , "Hmm..."
End Sub         'the reason I gave this button, is that I want to show that Chinese culture is not always as traditional as it was.
