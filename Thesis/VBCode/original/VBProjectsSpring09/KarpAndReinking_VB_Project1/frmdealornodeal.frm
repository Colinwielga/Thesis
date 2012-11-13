VERSION 5.00
Begin VB.Form frmdealornodeal 
   BackColor       =   &H00000000&
   Caption         =   "Deal or No Deal"
   ClientHeight    =   10305
   ClientLeft      =   2400
   ClientTop       =   255
   ClientWidth     =   11370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   11370
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   58
      Top             =   9600
      Width           =   1455
   End
   Begin VB.PictureBox piccasenumber 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   9000
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   31
      Top             =   9120
      Width           =   1335
   End
   Begin VB.CommandButton cmd26 
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   30
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd25 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmd24 
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   28
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd23 
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   27
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd22 
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   26
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd21 
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   25
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd20 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   24
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd18 
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd17 
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd16 
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd15 
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd14 
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd12 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd11 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd10 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picResultsfive 
      Height          =   855
      Left            =   7320
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsfour 
      Height          =   855
      Left            =   5640
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsthree 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   35
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwo 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsone 
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwelve 
      Height          =   855
      Left            =   9000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultseleven 
      Height          =   855
      Left            =   7320
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   42
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsten 
      Height          =   855
      Left            =   5640
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   41
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsnine 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   40
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultseight 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsseven 
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsthirteen 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   44
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsfourteen 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   45
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsfifteen 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultssixteen 
      Height          =   855
      Left            =   5640
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultsseventeen 
      Height          =   855
      Left            =   7320
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   48
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultseighteen 
      Height          =   855
      Left            =   9000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   49
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwenty 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   57
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentyone 
      Height          =   855
      Left            =   3960
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   51
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentytwo 
      Height          =   855
      Left            =   5640
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   52
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentythree 
      Height          =   855
      Left            =   7320
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   53
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentyfour 
      Height          =   855
      Left            =   9000
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   54
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentysix 
      Height          =   855
      Left            =   6360
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   56
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picResultstwentyfive 
      Height          =   855
      Left            =   3240
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   55
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd19 
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   23
      Top             =   6240
      Width           =   1215
   End
   Begin VB.PictureBox picResultsnineteen 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   50
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblthinking 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "<-- Click here if you want thinking music while deciding which case to pick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1095
      Left            =   3720
      TabIndex        =   60
      Top             =   9120
      Width           =   1935
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   2400
      OleObjectBlob   =   "frmdealornodeal.frx":0000
      SourceDoc       =   "M:\CS130\Project\thinking.mp3"
      TabIndex        =   59
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblpicked 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You Picked Case Number:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   735
      Left            =   7080
      TabIndex        =   32
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label lblno 
      BackColor       =   &H00000000&
      Caption         =   "NO DEAL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblr 
      BackColor       =   &H0000C0C0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblor 
      BackColor       =   &H0000C0C0&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lbldeal 
      BackColor       =   &H0000C0C0&
      Caption         =   "DEAL"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Deal or No Deal. Please pick a case number below to begin your game. Did you pick the million dollar case?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   9855
   End
End
Attribute VB_Name = "frmdealornodeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim First As Boolean
Dim I As Double
Dim Total As Single
'Project: Deal or No Deal
'frmdealornodeal
'Holly Reinking and Danielle Karp
'Written 3/15/09
'Purpose: Select a briefcase holding a certain amount of money (inputted from an array) and then dim that amount of money on frmmoney. Also, Compute the bank offer amount


Private Sub cmd1_Click()            'Holds an unknown amount of money to be revealed later in the game
      cmd1.Visible = False
      
      amount = CaseDollar(1)
      
       If First = False Then
            MsgBox "You picked case 1.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 1 is the first case they picked it is moved into storage
                piccasenumber.Print "1"
                First = True
                Num = 1
                Good = amount
        
        ElseIf First = True Then
            MsgBox "You picked case 1, inside is " & FormatCurrency(CaseDollar(1)) & ".", , "Case Number 1" 'If Case 1 is not the first case picked the user is told what amount of money that case held
            
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
            
            Select Case amount
                Case 0.01                                       'On the newly showing form the button holding the above amount of money is enabled as false
                    frmmoney.cmdmoney1.Enabled = False
                Case 1
                    frmmoney.cmdmoney2.Enabled = False
                Case 5
                    frmmoney.cmdmoney3.Enabled = False
                Case 10
                    frmmoney.cmdmoney4.Enabled = False
                Case 25
                    frmmoney.cmdmoney5.Enabled = False
                Case 50
                    frmmoney.cmdmoney6.Enabled = False
                Case 75
                    frmmoney.cmdmoney7.Enabled = False
                Case 100
                    frmmoney.cmdmoney8.Enabled = False
                Case 200
                    frmmoney.cmdmoney9.Enabled = False
                Case 300
                    frmmoney.cmdmoney10.Enabled = False
                Case 400
                    frmmoney.cmdmoney11.Enabled = False
                Case 500
                    frmmoney.cmdmoney12.Enabled = False
                Case 750
                    frmmoney.cmdmoney13.Enabled = False
                Case 1000
                    frmmoney.cmdmoney14.Enabled = False
                Case 5000
                    frmmoney.cmdmoney15.Enabled = False
                Case 10000
                    frmmoney.cmdmoney16.Enabled = False
                Case 25000
                    frmmoney.cmdmoney17.Enabled = False
                Case 50000
                    frmmoney.cmdmoney18.Enabled = False
                Case 75000
                    frmmoney.cmdmoney19.Enabled = False
                Case 100000
                    frmmoney.cmdmoney20.Enabled = False
                Case 200000
                    frmmoney.cmdmoney21.Enabled = False
                Case 300000
                    frmmoney.cmdmoney22.Enabled = False
                Case 400000
                    frmmoney.cmdmoney23.Enabled = False
                Case 500000
                    frmmoney.cmdmoney24.Enabled = False
                Case 750000
                    frmmoney.cmdmoney25.Enabled = False
                Case 1000000
                    frmmoney.cmdmoney26.Enabled = False
                End Select
   
    Sum = Sum - CaseDollar(1)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
    End If
    
End Sub

Private Sub cmd10_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd10.Visible = False
    
    amount = CaseDollar(10)
    
        If First = False Then
            MsgBox "You picked case 10.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 10 is the first case they picked it is moved into storage
                piccasenumber.Print "10"
                First = True
                Num = 10
                Good = amount
                
        ElseIf First = True Then
            MsgBox "You picked case 11, inside is " & FormatCurrency(CaseDollar(10)) & ".", , "Case Number 10" 'If Case 10 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                             'One form is hidden while another is shown
                frmmoney.Show

        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
 
    Sum = Sum - CaseDollar(10)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
  End If
            
End Sub

Private Sub cmd11_Click()           'Holds an unknown amount of money to be revealed later in the game
     cmd11.Visible = False
     
     amount = CaseDollar(11)
     
        If First = False Then
            MsgBox "You picked case 11.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If case 11 is the first case they pick it is moved into storage
                piccasenumber.Print "11"
                First = True
                Num = 11
                Good = amount
                
        ElseIf First = True Then
            MsgBox "You picked case 11, inside is " & FormatCurrency(CaseDollar(11)) & ".", , "Case Number 11" 'If Case 11 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
      
    Sum = Sum - CaseDollar(11)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd12_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd12.Visible = False
    
    amount = CaseDollar(12)
    
        If First = False Then
            MsgBox "You picked case 12.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 12 is the first case they picked it is moved into storage
                piccasenumber.Print "12"
                First = True
                Num = 12
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 12, inside is " & FormatCurrency(CaseDollar(12)) & ".", , "Case Number 12"  'If Case 12 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
       
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(12)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
            
End Sub

Private Sub cmd13_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd13.Visible = False
    
    amount = CaseDollar(13)
        If First = False Then
            MsgBox "You picked case 13.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 13 is the first case they picked it is moved into storage
                piccasenumber.Print "13"
                First = True
                Num = 13
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 13, inside is " & FormatCurrency(CaseDollar(13)) & ".", , "Case Number 13" 'If Case 13 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(13)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd14_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd14.Visible = False
    
    amount = CaseDollar(14)
    
        If First = False Then
            MsgBox "You picked case 14.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 14 is the first case they picked it is moved into storage
                piccasenumber.Print "14"
                First = True
                Num = 14
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 14, inside is " & FormatCurrency(CaseDollar(14)) & ".", , "Case Number 14" 'If Case 14 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
       
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
        
    Sum = Sum - CaseDollar(14)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
    
End Sub

Private Sub cmd15_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd15.Visible = False
    
    amount = CaseDollar(15)
    
        If First = False Then
            MsgBox "You picked case 15.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 15 is the first case they picked it is moved into storage
                piccasenumber.Print "15"
                First = True
                Num = 15
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 15, inside is " & FormatCurrency(CaseDollar(15)) & ".", , "Case Number 15" 'If Case 15 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
      
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
        
    Sum = Sum - CaseDollar(15)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
            
End Sub

Private Sub cmd16_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd16.Visible = False
    
    amount = CaseDollar(16)
    
        If First = False Then
            MsgBox "You picked case 16.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 16 is the first case they picked it is moved into storage
                piccasenumber.Print "16"
                First = True
                Num = 16
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 16, inside is " & FormatCurrency(CaseDollar(16)) & ".", , "Case Number 16" 'If Case 16 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(16)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd17_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd17.Visible = False
    
    amount = CaseDollar(17)
    
        If First = False Then
            MsgBox "You picked case 17.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 17 is the first case they picked it is moved into storage
                piccasenumber.Print "17"
                First = True
                Num = 17
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 17, inside is " & FormatCurrency(CaseDollar(17)) & ".", , "Case Number 17" 'If Case 17 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
      
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(17)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd18_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd18.Visible = False
    
    amount = CaseDollar(18)
    
        If First = False Then
            MsgBox "You picked case 18.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 18 is the first case they picked it is moved into storage
                piccasenumber.Print "18"
                First = True
                Num = 18
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 18, inside is " & FormatCurrency(CaseDollar(18)) & ".", , "Case Number 18" 'If Case 18 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(18)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd19_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd19.Visible = False
    
    amount = CaseDollar(19)
    
        If First = False Then
            MsgBox "You picked case 19.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 19 is the first case they picked it is moved into storage
                piccasenumber.Print "19"
                First = True
                Num = 19
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 19, inside is " & FormatCurrency(CaseDollar(19)) & ".", , "Case Number 19" 'If Case 19 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(19)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub


Private Sub cmd2_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd2.Visible = False
    
    amount = CaseDollar(2)
    
        If First = False Then
            MsgBox "You picked case 2.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 2 is the first case they picked it is moved into storage
                piccasenumber.Print "2"
                First = True
                Num = 2
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 2, inside is " & FormatCurrency(CaseDollar(2)) & ".", , "Case Number 2" 'If Case 2 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
       
    Sum = Sum - CaseDollar(2)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
    End If
    
End Sub

Private Sub cmd20_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd20.Visible = False
    
    amount = CaseDollar(20)
    
        If First = False Then
            MsgBox "You picked case 20.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 20 is the first case they picked it is moved into storage
                piccasenumber.Print "20"
                First = True
                Num = 20
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 20, inside is " & FormatCurrency(CaseDollar(20)) & ".", , "Case Number 20" 'If Case 20 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(20)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd21_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd21.Visible = False
    
    amount = CaseDollar(21)
    
        If First = False Then
            MsgBox "You picked case 21.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 21 is the first case they picked it is moved into storage
                piccasenumber.Print "21"
                First = True
                Num = 21
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 21, inside is " & FormatCurrency(CaseDollar(21)) & ".", , "Case Number 21" 'If Case 21 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
   
    Sum = Sum - CaseDollar(21)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd22_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd22.Visible = False
    
    amount = CaseDollar(22)
    
        If First = False Then
            MsgBox "You picked case 22.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 22 is the first case they picked it is moved into storage
                piccasenumber.Print "22"
                First = True
                Num = 22
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 22, inside is " & FormatCurrency(CaseDollar(22)) & ".", , "Case Number 22" 'If Case 22 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select

    Sum = Sum - CaseDollar(22)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
            
End Sub

Private Sub cmd23_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd23.Visible = False
    
    amount = CaseDollar(23)
    
        If First = False Then
            MsgBox "You picked case 23.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 23 is the first case they picked it is moved into storage
                piccasenumber.Print "23"
                First = True
                Num = 23
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 23, inside is " & FormatCurrency(CaseDollar(23)) & ".", , "Case Number 23" 'If Case 23 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
        
    Sum = Sum - CaseDollar(23)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd24_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd24.Visible = False
    
    amount = CaseDollar(24)
    
        If First = False Then
            MsgBox "You picked case 24.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 24 is the first case they picked it is moved into storage
                piccasenumber.Print "24"
                First = True
                Num = 24
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 24, inside is " & FormatCurrency(CaseDollar(24)) & ".", , "Case Number 24" 'If Case 24 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
        
    Sum = Sum - CaseDollar(24)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd25_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd25.Visible = False
    
    amount = CaseDollar(25)
    
        If First = False Then
            MsgBox "You picked case 25.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 25 is the first case they picked it is moved into storage
                piccasenumber.Print "25"
                First = True
                Num = 25
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 25, inside is " & FormatCurrency(CaseDollar(25)) & ".", , "Case Number 25" 'If Case 25 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
       
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
   
    Sum = Sum - CaseDollar(25)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd26_Click()           'Holds an unknown amount of money to be revealed later in the game
    cmd26.Visible = False
    
    amount = CaseDollar(26)
    
        If First = False Then
            MsgBox "You picked case 26.", , "This is your guess for the $1,000,000 case! Good Luck!"            'If Case 26 is the first case they picked it is moved into storage
                piccasenumber.Print "26"
                First = True
                Num = 26
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 26, inside is " & FormatCurrency(CaseDollar(26)) & ".", , "Case Number 26" 'If Case 26 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
        
    Sum = Sum - CaseDollar(26)          'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd3_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd3.Visible = False
    
    amount = CaseDollar(3)
    
        If First = False Then
            MsgBox "You picked case 3.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 3 is the first case they picked it is moved into storage
                piccasenumber.Print "3"
                First = True
                Num = 3
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 3, inside is " & FormatCurrency(CaseDollar(3)) & ".", , "Case Number 3" 'If Case 3 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
    
    Sum = Sum - CaseDollar(3)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd4_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd4.Visible = False
    
    amount = CaseDollar(4)
     
        If First = False Then
            MsgBox "You picked case 4.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 1 is the first case they picked it is moved into storage
                piccasenumber.Print "4"
                First = True
                Num = 4
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 4, inside is " & FormatCurrency(CaseDollar(4)) & ".", , "Case Number 4" 'If Case 4 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
       
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
 
    Sum = Sum - CaseDollar(4)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
            
End Sub

Private Sub cmd5_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd5.Visible = False
    
    amount = CaseDollar(5)
     
        If First = False Then
            MsgBox "You picked case 5.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 5 is the first case they picked it is moved into storage
                piccasenumber.Print "5"
                First = True
                Num = 5
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 5, inside is " & FormatCurrency(CaseDollar(5)) & ".", , "Case Number 5" 'If Case 5 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
       
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
       
    Sum = Sum - CaseDollar(5)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd6_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd6.Visible = False
    
    amount = CaseDollar(6)
    
        If First = False Then
            MsgBox "You picked case 6.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 6 is the first case they picked it is moved into storage
                piccasenumber.Print "6"
                First = True
                Num = 6
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 6, inside is " & FormatCurrency(CaseDollar(6)) & ".", , "Case Number 6" 'If Case 6 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
      
    Sum = Sum - CaseDollar(6)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd7_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd7.Visible = False
    
    amount = CaseDollar(7)
    
        If First = False Then
            MsgBox "You picked case 7.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 7 is the first case they picked it is moved into storage
                piccasenumber.Print "7"
                First = True
                Num = 7
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 7, inside is " & FormatCurrency(CaseDollar(7)) & ".", , "Case Number 7" 'If Case 7 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
  
    Sum = Sum - CaseDollar(7)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
      End If
      
End Sub

Private Sub cmd8_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd8.Visible = False
    
    amount = CaseDollar(8)
    
        If First = False Then
            MsgBox "You picked case 8.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 8 is the first case they picked it is moved into storage
                piccasenumber.Print "8"
                First = True
                Num = 8
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 8, inside is " & FormatCurrency(CaseDollar(8)) & ".", , "Case Number 8" 'If Case 8 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
      
        
    Sum = Sum - CaseDollar(8)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
        End If
 
End Sub

Private Sub cmd9_Click()            'Holds an unknown amount of money to be revealed later in the game
    cmd9.Visible = False
    
    amount = CaseDollar(9)
    
        If First = False Then
            MsgBox "You picked case 9.", , "This is your guess for the $1,000,000 case! Good Luck!"         'If Case 9 is the first case they picked it is moved into storage
                piccasenumber.Print "9"
                First = True
                Num = 9
                Good = amount
        ElseIf First = True Then
            MsgBox "You picked case 9, inside is " & FormatCurrency(CaseDollar(9)) & ".", , "Case Number 9" 'If Case 9 is not the first case picked the user is told what amount of money that case held
                
                frmdealornodeal.Hide                            'One form is hidden while another is shown
                frmmoney.Show
        
        Select Case amount
            Case 0.01                                           'On the newly showing form the button holding the above amount of money is enabled as false
                frmmoney.cmdmoney1.Enabled = False
            Case 1
                frmmoney.cmdmoney2.Enabled = False
            Case 5
                frmmoney.cmdmoney3.Enabled = False
            Case 10
                frmmoney.cmdmoney4.Enabled = False
            Case 25
                frmmoney.cmdmoney5.Enabled = False
            Case 50
                frmmoney.cmdmoney6.Enabled = False
            Case 75
                frmmoney.cmdmoney7.Enabled = False
            Case 100
                frmmoney.cmdmoney8.Enabled = False
            Case 200
                frmmoney.cmdmoney9.Enabled = False
            Case 300
                frmmoney.cmdmoney10.Enabled = False
            Case 400
                frmmoney.cmdmoney11.Enabled = False
            Case 500
                frmmoney.cmdmoney12.Enabled = False
            Case 750
                frmmoney.cmdmoney13.Enabled = False
            Case 1000
                frmmoney.cmdmoney14.Enabled = False
            Case 5000
                frmmoney.cmdmoney15.Enabled = False
            Case 10000
                frmmoney.cmdmoney16.Enabled = False
            Case 25000
                frmmoney.cmdmoney17.Enabled = False
            Case 50000
                frmmoney.cmdmoney18.Enabled = False
            Case 75000
                frmmoney.cmdmoney19.Enabled = False
            Case 100000
                frmmoney.cmdmoney20.Enabled = False
            Case 200000
                frmmoney.cmdmoney21.Enabled = False
            Case 300000
                frmmoney.cmdmoney22.Enabled = False
            Case 400000
                frmmoney.cmdmoney23.Enabled = False
            Case 500000
                frmmoney.cmdmoney24.Enabled = False
            Case 750000
                frmmoney.cmdmoney25.Enabled = False
            Case 1000000
                frmmoney.cmdmoney26.Enabled = False
            End Select
   
    Sum = Sum - CaseDollar(9)           'To keep track of the sum by subracting out the amount of money in this case
    K = K + 1                           'To add one to a counter so the computer knows how to compute certain functions and dim the "return to cases" button
    
        End If
        
End Sub
    
Private Sub cmdQuit_Click()                     'To Quit the program
    End
End Sub



