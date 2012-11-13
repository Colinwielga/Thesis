VERSION 5.00
Begin VB.Form frmgreatwall 
   Caption         =   "Form1"
   ClientHeight    =   12825
   ClientLeft      =   6870
   ClientTop       =   765
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   Picture         =   "frmgreatwall.frx":0000
   ScaleHeight     =   12825
   ScaleWidth      =   14640
   Begin VB.CommandButton Command13 
      Caption         =   "Homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8400
      TabIndex        =   13
      Top             =   10440
      Width           =   5055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11520
      TabIndex        =   12
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      TabIndex        =   11
      Top             =   8880
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   7680
      ScaleHeight     =   7995
      ScaleWidth      =   6435
      TabIndex        =   10
      Top             =   360
      Width           =   6495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Water Cube"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   9
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Beijing Park"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   8
      Top             =   11760
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "The Winter Palace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   9000
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Summer Palace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   9000
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "The Palace Museum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tianan Square"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Beihai Park"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GreatWall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bird's Nest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Temple of Heaven"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "frmgreatwall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runningtotal As Integer


Private Sub Command1_Click()
runningtotal = runningtotal + 10
picResults.Print "Temple of Heaven"; Tab(30); "$10"
End Sub

Private Sub Command10_Click()
runningtotal = runningtotal + 10
picResults.Print "Water Cube"; Tab(30); "$10"
End Sub

Private Sub Command11_Click()
Dim total As Integer
total = runningtotal
picResults.Print "-------------------------------------------------------"
picResults.Print "Total:", Tab(30); FormatCurrency(total)
End Sub

Private Sub Command12_Click()
runningtotal = 0
picResults.Cls
End Sub

Private Sub Command13_Click()
frmgreatwall.Hide
frmMain.Show

End Sub

Private Sub Command2_Click()
runningtotal = runningtotal + 20
picResults.Print "Bird's Nest"; Tab(30); "$20"
End Sub

Private Sub Command3_Click()
runningtotal = runningtotal + 10
picResults.Print "GreatWall"; Tab(30); "$10"
End Sub

Private Sub Command4_Click()
runningtotal = runningtotal + 30
picResults.Print "Beihai Park"; Tab(30); "$30"

End Sub

Private Sub Command5_Click()
runningtotal = runningtotal + 0
picResults.Print "Tianan Square"; Tab(30); "Free"
End Sub

Private Sub Command6_Click()
runningtotal = runningtotal + 5
picResults.Print "The Palace Museum"; Tab(30); "$5"
End Sub

Private Sub Command7_Click()
runningtotal = runningtotal + 10
picResults.Print "Summer Palace"; Tab(30); "$10"
End Sub

Private Sub Command8_Click()
runningtotal = runningtotal + 30
picResults.Print "The Winter Palace"; Tab(30); "$30"
End Sub

Private Sub Command9_Click()
runningtotal = runningtotal + 6
picResults.Print "Beijing Park"; Tab(30); "$6"
End Sub

