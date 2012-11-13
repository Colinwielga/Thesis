VERSION 5.00
Begin VB.Form frmsinidol 
   BackColor       =   &H8000000A&
   Caption         =   "Singapore Idol"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
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
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   735
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
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "See the result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdresult1 
      Caption         =   "See the result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "See the result"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblidol 
      BackColor       =   &H8000000D&
      Caption         =   "Which is the Winner of Singapore Idol??                "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Image Image3 
      Height          =   3840
      Left            =   4920
      Picture         =   "frmsinidol.frx":0000
      Top             =   120
      Width           =   2160
   End
   Begin VB.Image Image2 
      Height          =   3090
      Left            =   2520
      Picture         =   "frmsinidol.frx":1B042
      Top             =   480
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   0
      Picture         =   "frmsinidol.frx":33FA4
      Top             =   480
      Width           =   2580
   End
End
Attribute VB_Name = "frmsinidol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Tokyo, Berlin, Singapore- My Summer 2005 (Makihara_Kosuke.vbp)
'Form Name: Singapore Idol(frmsinidol.frm)
'Author: Kosuke Makihara
'Date Wrriten: 27 Oct 2005
'Ojectives:
'This form ask user to choose one picture which look like "Singapore Idol", and
'each button show the result of choice by the messege box.



Private Sub cmdback_Click()
frmsinidol.Hide
frmsingapore.Show

End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub cmdresult1_Click()
'Thie code works to show the result of user's choice by messege box
MsgBox "You know who she is... she isn't a Singaporean...", , " Result"

End Sub

Private Sub Command1_Click()
'Thie code works to show the result of user's choice by messege box
MsgBox "He is a Prime Minister of Singapore", , "Result"


End Sub

Private Sub Command3_Click()
'Thie code works to show the result of user's choice by messege box
MsgBox "Yes... He, Muhammad Taufik bin Batisah, is the winner of Singapore Idol in the begining of 2005. His ethnic origin is Malay, as his name indicates.", , "Result"


End Sub

