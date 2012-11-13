VERSION 5.00
Begin VB.Form frmdream 
   BackColor       =   &H0000C000&
   Caption         =   "Dream Team"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form4"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtfwd1 
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtfwd2 
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtmidfwd 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtmid 
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtwgrgt 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtwgleft 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtgoalie 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtdefleft 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtcentral2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtcentral1 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtdefright 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Main Menu"
      Height          =   615
      Left            =   7440
      TabIndex        =   1
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   855
      Left            =   4680
      Top             =   4440
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   4920
      X2              =   4920
      Y1              =   6600
      Y2              =   3240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   1335
      Left            =   8280
      Top             =   4200
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      FillColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   720
      Top             =   4200
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9120
      X2              =   9120
      Y1              =   6600
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   720
      X2              =   9120
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   720
      X2              =   720
      Y1              =   3240
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   720
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label ldlDreamteam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DREAM TEAM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmdream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMenu_Click()
frmChampions.Show
frmdream.Hide
End Sub

