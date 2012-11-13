VERSION 5.00
Begin VB.Form others
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   Picture         =   "others.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1
      BackColor       =   &H0080FF80&
      Caption         =   "Return"
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Timer timeslide
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   6480
   End
   Begin VB.CommandButton cmdslide
      BackColor       =   &H0000FFFF&
      Caption         =   "Slide show!!"
      BeginProperty Font
         Name            =   "Mathematica6"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   3015
   End
   Begin VB.PictureBox picResults
      AutoSize        =   -1  'True
      Height          =   4575
      Left            =   2280
      ScaleHeight     =   4515
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label labshow
      BackColor       =   &H0080C0FF&
      Caption         =   "Let's watch a slide show from other area's food in China"
      BeginProperty Font
         Name            =   "Magneto"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "others"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Chinese food
'Form Name: other
'Author: Yifan Pang
'Date Written: feb 23 2010
'The purpose of this form s a slide show to show other food in China
Option Explicit
Dim I As Integer
Dim pic As Integer
Dim food(1 To 50) As String
Dim allthingsthataregood As Long

Private Sub cmdslide_Click() 'call the time to work
pic = 1
timeslide.Enabled = True
End Sub

Private Sub Command1_Click()
others.Hide
China.Show
End Sub

Private Sub Form_Load() 'auto load the picture

I = 0
Open App.Path & "\picture.txt" For Input As #1
Do While Not EOF(1)
    I = I + 1
    Input #1, food(I)
Loop
Close #1
End Sub

Private Sub timeslide_Timer() 'timer for slide show
    If pic < 8 Then
        picResults.picture = LoadPicture(App.Path & "\" & food(pic))
        pic = pic + 1
    Else
        timeslide.Enabled = False
    End If
End Sub

