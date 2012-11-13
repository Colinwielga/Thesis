VERSION 5.00
Begin VB.Form Minnesota 
   BackColor       =   &H00800000&
   Caption         =   "Form2"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   FillColor       =   &H00C0FFFF&
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFFF&
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMarshall 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   255
   End
   Begin VB.CommandButton cmdBrainerd 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdDuluth 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5760
      Picture         =   "Form2.frx":0BB1
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      Picture         =   "Form2.frx":C083
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6360
      Picture         =   "Form2.frx":14BB5
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdStPaul 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label lblMarshall 
      BackColor       =   &H00008000&
      Caption         =   "Marshall"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblDuluth 
      BackColor       =   &H00008000&
      Caption         =   "Duluth"
      BeginProperty Font 
         Name            =   "Minion Pro Med"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblBrainerd 
      BackColor       =   &H00008000&
      Caption         =   "Brainerd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblstpaul 
      BackColor       =   &H00008000&
      Caption         =   "St. Paul"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Select a city in Minnesota!"
      BeginProperty Font 
         Name            =   "Myriad Web Pro Condensed"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Minnesota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrainerd_Click()
Form1.Show
Minnesota.Hide

End Sub

Private Sub cmdDuluth_Click()
Duluth.Show
Minnesota.Hide

End Sub

Private Sub cmdMarshall_Click()
Minnesota.Hide
Marshall.Show
End Sub
