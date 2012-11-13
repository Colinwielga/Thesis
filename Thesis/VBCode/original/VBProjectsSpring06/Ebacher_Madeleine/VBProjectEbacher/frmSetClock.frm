VERSION 5.00
Begin VB.Form frmSetClock 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Step 1: twenty four or twelve?"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   2400
   End
   Begin VB.PictureBox picOutput 
      Height          =   855
      Left            =   240
      ScaleHeight     =   795
      ScaleWidth      =   2115
      TabIndex        =   15
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdTimeShow 
      Caption         =   "Show"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtNowSecond 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   12
      Text            =   "00"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtNowMinute 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   10
      Text            =   "00"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtNowHour 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Text            =   "00"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSetClock 
      BackColor       =   &H00FF8080&
      Caption         =   "Set Clock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblMyProgram 
      BackColor       =   &H00FFC0C0&
      Caption         =   "VB Alarm Clock - Madeleine Ebacher"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Line Line10 
      X1              =   480
      X2              =   480
      Y1              =   1200
      Y2              =   720
   End
   Begin VB.Line Line9 
      X1              =   2640
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line8 
      X1              =   2280
      X2              =   2280
      Y1              =   840
      Y2              =   1440
   End
   Begin VB.Label lblColon2 
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   11
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblColon 
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblCurrentIs 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The current time is:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblCurrent 
      BackColor       =   &H00FFC0C0&
      Caption         =   "current time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblSecond 
      Caption         =   "second"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblMinute 
      Caption         =   "minute"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblHour 
      Caption         =   "hour"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Line Line7 
      X1              =   1920
      X2              =   1920
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   4320
      X2              =   4320
      Y1              =   720
      Y2              =   1200
   End
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   2640
      Y1              =   840
      Y2              =   1200
   End
   Begin VB.Label lblFormat 
      BackColor       =   &H00FFC0C0&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "time is displayed in the format:          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSetClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NowHour, NowMinute, NowSecond As Integer
Dim TwentyFourTime(0 To 24) As Integer
Dim TwelveTime(1 To 12) As Integer

Private Sub cmdSetClock_Click()
    frmSetClock.Hide
    frmEnglishMenu.Show
End Sub

Private Sub cmdTimeShow_Click()
    txtNowHour.Text = DateTime.Now.Hour
    txtNowMinute.Text = DateTime.Now.Minute
    txtNowSecond.Text = DateTime.Now.Second
End Sub

Private Sub Picture1_Click()
    picOutput.Print DateTime
    
End Sub

Private Sub txtNowHour_Change()
    txtNowHour.Text = DateTime.Now.Hour
End Sub
