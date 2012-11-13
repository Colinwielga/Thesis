VERSION 5.00
Begin VB.Form frmMountains 
   BackColor       =   &H00FFFF00&
   Caption         =   "Mountains"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   Picture         =   "frmMountains.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext4 
      Caption         =   "Next!"
      Height          =   615
      Left            =   4800
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdNext3 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   5040
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton opt9 
      Caption         =   "Pikachu"
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton opt8 
      Caption         =   "A Pony"
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton opt7 
      Caption         =   "Charmander"
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdNext2 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton opt6 
      Caption         =   "Fifteen Million"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton opt5 
      Caption         =   "No one will ever know"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton opt4 
      Caption         =   "Three"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Question"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton opt3 
      Caption         =   "Patrick"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Ash"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Squidward"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdno 
      Caption         =   "No Way Jose!"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdyes 
      Caption         =   "Bring it on!"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "You made it passed the Mountain Man!!!!  Time to go to the Forest!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lbl5 
      Caption         =   "Who was Ash Ketchum's first pokemon?"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "How many licks does it take to get to the center of a tootsie pop?"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Caption         =   "Who is Spongebob's Best Friend?"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmMountains.frx":9726
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image img2 
      Height          =   4320
      Left            =   120
      Picture         =   "frmMountains.frx":97BE
      Top             =   120
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Do you dare to cross these mountains????"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   9
      Left            =   6240
      Picture         =   "frmMountains.frx":1098A
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   8
      Left            =   0
      Picture         =   "frmMountains.frx":11555
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   7
      Left            =   3960
      Picture         =   "frmMountains.frx":12120
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   6
      Left            =   1800
      Picture         =   "frmMountains.frx":12CEB
      Top             =   0
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   5
      Left            =   2760
      Picture         =   "frmMountains.frx":138B6
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   4
      Left            =   3840
      Picture         =   "frmMountains.frx":14481
      Top             =   0
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   3
      Left            =   4920
      Picture         =   "frmMountains.frx":1504C
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   2
      Left            =   5880
      Picture         =   "frmMountains.frx":15C17
      Top             =   0
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   1
      Left            =   1800
      Picture         =   "frmMountains.frx":167E2
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   1665
      Index           =   0
      Left            =   480
      Picture         =   "frmMountains.frx":173AD
      Top             =   1560
      Width           =   1740
   End
End
Attribute VB_Name = "frmMountains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, the user has two lives and has to go through three
'trivia questions. If they get two  wrong they will die.
'
Dim Lives As Integer
Private Sub CmdNext_Click()
 If opt3.Value = True Then
    MsgBox "Good Job! You're right!", , "Right"
    Lbl3.Visible = False
    Opt1.Visible = False
    opt2.Visible = False
    opt3.Visible = False
    cmdNext.Visible = False
    lbl4.Visible = True
    opt4.Visible = True
    opt5.Visible = True
    opt6.Visible = True
    cmdNext2.Visible = True
    
 Else
    Lives = Lives - 1
    If Lives = 0 Then
        MsgBox "You are out of lives. The Mountain Man is going to eat you now", , "You are being eaten"
        MsgBox "This is where your story ends. Start Over", , "Story Ends"
        frmMountains.Hide
        frmWelcome.Show
         lbl1.Visible = True
        cmdyes.Visible = True
        cmdno.Visible = True
        img2.Visible = False
        lbl2.Visible = False
        Lbl3.Visible = False
        Opt1.Visible = False
        opt2.Visible = False
        opt3.Visible = False
        cmdNext.Visible = False
        Lives = 2
    Else
        MsgBox "Sorry, that isn't quite right.  The Mountain Man was looking for Patrick as an answer.  Careful! You only have " & Lives & " left!", , "Wrong"
        Lbl3.Visible = False
    Opt1.Visible = False
    opt2.Visible = False
    opt3.Visible = False
    cmdNext.Visible = False
    lbl4.Visible = True
    opt4.Visible = True
    opt5.Visible = True
    opt6.Visible = True
    cmdNext2.Visible = True
        
    End If
    End If
    
End Sub

Private Sub cmdNext2_Click()
    If opt5.Value = True Then
    MsgBox "Good Job! You're right!", , "Right"
    lbl4.Visible = False
    opt4.Visible = False
    opt5.Visible = False
    opt6.Visible = False
    cmdNext2.Visible = False
    opt7.Visible = True
    opt8.Visible = True
    opt9.Visible = True
    lbl5.Visible = True
    cmdNext3.Visible = True
 Else
    Lives = Lives - 1
    If Lives = 0 Then
        MsgBox "You are out of lives. The Mountain Man is going to eat you now", , "You are being eaten"
        MsgBox "This is where your story ends. Start Over", , "Story Ends"
        frmMountains.Hide
        frmWelcome.Show
         lbl1.Visible = True
        cmdyes.Visible = True
        cmdno.Visible = True
        img2.Visible = False
        lbl2.Visible = False
        lbl4.Visible = False
        opt4.Visible = False
        opt5.Visible = False
        opt6.Visible = False
        cmdNext2.Visible = False
        Lives = 2
    Else
        MsgBox "Sorry, that isn't quite right.  The Mountain Man was looking for No one will never know as an answer.  Careful! You only have " & Lives & " left!", , "Wrong"
        lbl4.Visible = False
    opt4.Visible = False
    opt5.Visible = False
    opt6.Visible = False
    cmdNext2.Visible = False
    opt7.Visible = True
    opt8.Visible = True
    opt9.Visible = True
    lbl5.Visible = True
    cmdNext3.Visible = True
    End If
    End If
    
    
End Sub

Private Sub cmdNext3_Click()
    If opt9.Value = True Then
    MsgBox "Good Job! You're right!", , "Right"
        lbl2.Visible = False
    lbl5.Visible = False
    opt7.Visible = False
    opt8.Visible = False
    opt9.Visible = False
    cmdNext3.Visible = False
    Lbl6.Visible = True
    cmdNext4.Visible = True
 Else
    Lives = Lives - 1
    If Lives = 0 Then
        MsgBox "You are out of lives. The Mountain Man is going to eat you now", , "You are being eaten"
        MsgBox "This is where your story ends. Start Over", , "Story Ends"
        frmMountains.Hide
        frmWelcome.Show
         lbl1.Visible = True
        cmdyes.Visible = True
        cmdno.Visible = True
        img2.Visible = False
        lbl2.Visible = False
        lbl5.Visible = False
        opt7.Visible = False
        opt8.Visible = False
        opt9.Visible = False
        cmdNext3.Visible = False
        Lives = 2
    Else
        MsgBox "Sorry, that isn't quite right.  The Mountain Man was looking for Pikachu as an answer.  Careful! You only have " & Lives & " left!", , "Wrong"
        lbl2.Visible = False
    lbl5.Visible = False
    opt7.Visible = False
    opt8.Visible = False
    opt9.Visible = False
    cmdNext3.Visible = False
    Lbl6.Visible = True
    cmdNext4.Visible = True
    End If
    End If

End Sub

Private Sub cmdNext4_Click()
    lbl1.Visible = True
    cmdyes.Visible = True
    cmdno.Visible = True
    img2.Visible = False
    Lbl6.Visible = False
    cmdNext4.Visible = False
    frmMountains.Hide
    frmForest.Show
End Sub

Private Sub cmdno_Click()
    MsgBox "A wizard never turns down an adventure!! Your powers have been taken away! You will live as a mortal forever!", , "Bad Choice"
    MsgBox "This is where your story ends. Start Over", , "Story Ends."
    frmMountains.Hide
    frmWelcome.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdyes_Click()
    lbl1.Visible = False
    cmdyes.Visible = False
    cmdno.Visible = False
    img2.Visible = True
    lbl2.Visible = True
    Lbl3.Visible = True
    Opt1.Visible = True
    opt2.Visible = True
    opt3.Visible = True
    cmdNext.Visible = True
    Lives = 2
End Sub
