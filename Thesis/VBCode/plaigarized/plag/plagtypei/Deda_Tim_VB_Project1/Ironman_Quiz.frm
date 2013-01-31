VERSION 5.00
Begin VB.Form Ironman_Quiz
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   5280
   ClientTop       =   3315
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   9930
   Begin VB.CommandButton Results
      BackColor       =   &H0000C0C0&
      Caption         =   "See Results"
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton Tony2
      BackColor       =   &H0000C0C0&
      Caption         =   "Tony Stark"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Everhart2
      BackColor       =   &H0000C0C0&
      Caption         =   "Christine Everhart"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Pepper2
      BackColor       =   &H0000C0C0&
      Caption         =   "Pepper Potts"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Stane
      BackColor       =   &H0000C0C0&
      Caption         =   "Obadiah Stane"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Will
      BackColor       =   &H0000C0C0&
      Caption         =   "William Ginter Riva"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Rhodey2
      BackColor       =   &H0000C0C0&
      Caption         =   "Rhodey"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Phil2
      BackColor       =   &H0000C0C0&
      Caption         =   "Agent Phil Coulson"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Pepper
      BackColor       =   &H0000C0C0&
      Caption         =   "Pepper Potts"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Yinsen
      BackColor       =   &H0000C0C0&
      Caption         =   "Yinsen"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Allen
      BackColor       =   &H0000C0C0&
      Caption         =   "Major Allen"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Everhart
      BackColor       =   &H0000C0C0&
      Caption         =   "Christine Everhart"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Phil
      BackColor       =   &H0000C0C0&
      Caption         =   "Agent Phil Coulson"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox PicResults
      BackColor       =   &H000000C0&
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   360
      ScaleHeight     =   1995
      ScaleWidth      =   4755
      TabIndex        =   8
      Top             =   7080
      Width           =   4815
   End
   Begin VB.CommandButton Gabriel
      BackColor       =   &H0000C0C0&
      Caption         =   "General Gabriel"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Tony
      BackColor       =   &H0000C0C0&
      Caption         =   "Tony Stark"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Rhodey
      BackColor       =   &H0000C0C0&
      Caption         =   "Rhodey"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Hogan
      BackColor       =   &H0000C0C0&
      Caption         =   "Hogan"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton IronmanReturn
      BackColor       =   &H0000C0C0&
      Caption         =   "Return to Ironman"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton MainReturn
      BackColor       =   &H0000C0C0&
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label6
      BackStyle       =   0  'Transparent
      Caption         =   $"Ironman_Quiz.frx":0000
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   600
      TabIndex        =   20
      Top             =   4680
      Width           =   8895
   End
   Begin VB.Label Label5
      BackStyle       =   0  'Transparent
      Caption         =   """For three hours. For three hours you got me standing here!"""
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   3360
      Width           =   7695
   End
   Begin VB.Label Label4
      BackStyle       =   0  'Transparent
      Caption         =   """Just call us S.H.I.E.L.D."""
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label3
      BackStyle       =   0  'Transparent
      Caption         =   "Results"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label2
      BackStyle       =   0  'Transparent
      Caption         =   """Well let's face it. This is not the worst thing you've caught me doing."""
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   8895
   End
   Begin VB.Label Label1
      BackStyle       =   0  'Transparent
      Caption         =   "Who Said That Line?"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1
      Height          =   9600
      Left            =   0
      Picture         =   "Ironman_Quiz.frx":009A
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Ironman_Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Found1 As Boolean, Found2 As Boolean, Found3 As Boolean, Found4 As Boolean


Private Sub Allen_Click()
    Everhart.Visible = False
    Phil.Visible = False
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Allen.Visible = True
    Yinsen.Visible = False
End Sub

Private Sub Everhart_Click()
    Everhart.Visible = True
    Phil.Visible = False
    Allen.Visible = False
    Yinsen.Visible = False
End Sub

Private Sub Everhart2_Click()
    Pepper2.Visible = False
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Stane.Visible = False
    Everhart2.Visible = True
    Tony2.Visible = False
End Sub

Private Sub Form_Load()
Found1 = False
Found2 = False
Found3 = False
Found4 = False
End Sub

Private Sub Gabriel_Click()
    Hogan.Visible = False
    Rhodey.Visible = False
    Tony.Visible = False
    Gabriel.Visible = True
End Sub

Private Sub Hogan_Click()
    Hogan.Visible = True
    Rhodey.Visible = False
    Tony.Visible = False
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Gabriel.Visible = False
End Sub

Private Sub IronmanReturn_Click()
Ironman.Show
Ironman_Quiz.Hide
End Sub

    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
Private Sub MainReturn_Click()
MainMenu.Show
Ironman_Quiz.Hide
End Sub

Private Sub Pepper_Click()
    Will.Visible = False
    Pepper.Visible = True
    Phil2.Visible = False
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Rhodey2.Visible = False
End Sub

Private Sub Pepper2_Click()
    Pepper2.Visible = True
    Stane.Visible = False
    Everhart2.Visible = False
    Tony2.Visible = False
    Found4 = True
End Sub

    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
Private Sub Phil_Click()
    Everhart.Visible = False
    Phil.Visible = True
    Allen.Visible = False
    Yinsen.Visible = False
    Found2 = True
End Sub

Private Sub Phil2_Click()
    Will.Visible = False
    Pepper.Visible = False
    Phil2.Visible = True
    Rhodey2.Visible = False
End Sub

Private Sub Results_Click()
If Found1 = True Then
    PicResults.Print "you got the first one right!"
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
Else
    PicResults.Print "The correct answer for the first one was Tony Stark"
End If
If Found2 = True Then
    PicResults.Print "you got the second one right!"
Else
    PicResults.Print "The correct answer for the second one was Agent Phil Coulson"
End If
If Found3 = True Then
    PicResults.Print "you got the third one right!"
Else
    PicResults.Print "The correct answer for the third one was Rhodey"
End If
If Found4 = True Then
    PicResults.Print "you got the last one right!"
Else
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    PicResults.Print "The correct answer for the last one was Pepper Potts"
End If
If Found1 = True And Found2 = True And Found3 = True And Found4 = True Then
    PicResults.Print "Wow! you got them all correct! Great job "; UserName
End If
End Sub

Private Sub Rhodey_Click()
    Hogan.Visible = False
    Rhodey.Visible = True
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Tony.Visible = False
    Gabriel.Visible = False
End Sub

Private Sub Rhodey2_Click()
    Will.Visible = False
    Pepper.Visible = False
    Phil2.Visible = False
    Rhodey2.Visible = True
    Found3 = True
End Sub

Private Sub Stane_Click()
    Pepper2.Visible = False
    Stane.Visible = True
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Everhart2.Visible = False
    Tony2.Visible = False
End Sub

Private Sub Tony_Click()
    Hogan.Visible = False
    Rhodey.Visible = False
    Tony.Visible = True
    Gabriel.Visible = False
    Found1 = True
End Sub

Private Sub Tony2_Click()
    Pepper2.Visible = False
    Stane.Visible = False
    Everhart2.Visible = False
    Tony2.Visible = True
End Sub

Private Sub Will_Click()
    Will.Visible = True
    Pepper.Visible = False
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six
    Phil2.Visible = False
    Rhodey2.Visible = False
End Sub
    ' a few big blocks of comments
    ' here, there, and everywhere
    ' words words words
    ' and here and stuff
    ' a few more lines
    ' and we should be good
    ' one one two two six

Private Sub Yinsen_Click()
    Everhart.Visible = False
    Phil.Visible = False
    Allen.Visible = False
    Yinsen.Visible = True
End Sub
