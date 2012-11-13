VERSION 5.00
Begin VB.Form frmMoney 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Your Cash Winnings"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEndingTime 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox txtGrandTotal 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Text            =   " "
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtPlayerMoney 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3240
      TabIndex        =   8
      Text            =   " "
      Top             =   1380
      Width           =   3975
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Memo:"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   8520
      X2              =   10560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1680
      X2              =   8160
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblDate 
      BackColor       =   &H000080FF&
      Caption         =   "Friday November 3, 2006"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   7080
      X2              =   10320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblJsig 
      BackColor       =   &H00000000&
      Caption         =   "Jeopardy Game Show"
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7080
      TabIndex        =   5
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblSignature 
      BackColor       =   &H00000000&
      Caption         =   "Signature"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblCheckNumber 
      BackColor       =   &H00000000&
      Caption         =   "1 5647 3245 7571 2378"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label lblHollywood 
      BackColor       =   &H000080FF&
      Caption         =   "Hollywood, California"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label lblOwner 
      BackColor       =   &H000080FF&
      Caption         =   "Jeopardy Game Show"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblCheckTo 
      BackColor       =   &H000080FF&
      Caption         =   "Pay To The Order Of"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4305
      Left            =   -120
      Picture         =   "frmMoney.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11505
   End
End
Attribute VB_Name = "frmMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Player's Check: This is the check the player will receive and it
'                will show the amount the user won.The Exit button
'                is the button to end the game, the user will not
'                be able to begin a new game without reopening the
'                program

Private Sub cmdExit_Click()
    End
End Sub

