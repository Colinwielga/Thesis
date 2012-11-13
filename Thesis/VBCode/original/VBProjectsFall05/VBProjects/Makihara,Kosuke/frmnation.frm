VERSION 5.00
Begin VB.Form frmnation 
   Caption         =   "Prime Minister "
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdpic2 
      Caption         =   "See the Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdpic3 
      Caption         =   "See the Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdpic4 
      Caption         =   "See the Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdpic1 
      Caption         =   "See the Result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Which is Japan's Prime Minister??"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
   End
   Begin VB.Image Image1 
      Height          =   3630
      Left            =   0
      Picture         =   "frmnation.frx":0000
      Top             =   360
      Width           =   3300
   End
   Begin VB.Image Image4 
      Height          =   2760
      Left            =   6000
      Picture         =   "frmnation.frx":2702A
      Top             =   600
      Width           =   3900
   End
   Begin VB.Image Image3 
      Height          =   3165
      Left            =   2760
      Picture         =   "frmnation.frx":4A10C
      Top             =   600
      Width           =   3300
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   9720
      Picture         =   "frmnation.frx":6C14A
      Top             =   600
      Width           =   2580
   End
End
Attribute VB_Name = "frmnation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Prime Minister(frmnation.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'This form ask user to choose one picture which look like Japan's prime minister, and
'each button show the result of choice by the messege box.

Private Sub cmdback_Click()
'back to main page of Tokyo
frmnation.Hide
frmtokyo.Show

End Sub

Private Sub cmdpic1_Click()
'Thie code works to show the result of choice by messege box
MsgBox "Yes, right side is Prime Minister Koizumi, and left side is President Bush...They are good buddy...", , "Result"

End Sub

Private Sub cmdpic2_Click()
'Thie code works to show the result of choice by messege box
MsgBox "No...He is Hu Jintao, a head of Chinese Communist Party", , "Result"


End Sub

Private Sub cmdpic3_Click()
'Thie code works to show the result of choice by messege box
MsgBox "No... left side is President Bush and right side is... he is currently famous dictator of North Korea, Kim Jungill working at Wendies", , "Result"
End Sub

Private Sub cmdpic4_Click()
'Thie code works to show the result of choice by messege box
MsgBox "Yes, left side is Prime Minister Koizumi, and right side is Richard Gere who came to Japan for his movie's PR", , "Result"

End Sub

Private Sub cmdquit_Click()
End

End Sub
