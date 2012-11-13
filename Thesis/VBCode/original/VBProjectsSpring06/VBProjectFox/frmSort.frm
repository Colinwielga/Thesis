VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H00004000&
   Caption         =   "Sorting your Trees"
   ClientHeight    =   6315
   ClientLeft      =   2310
   ClientTop       =   2550
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   Picture         =   "frmSort.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   9945
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back to Beginning "
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort them by type"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Them into Deciduous and Evergreens"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "By Kelly Fox"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "How do you want to sort the trees in your file?"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   7575
   End
   Begin VB.Image imgBackSort 
      Height          =   9000
      Index           =   1
      Left            =   0
      Picture         =   "frmSort.frx":0342
      Top             =   -120
      Width           =   12000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   5880
      Index           =   0
      Left            =   2640
      Picture         =   "frmSort.frx":15FC84
      Top             =   840
      Width           =   4500
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00004000&
      Caption         =   "How do you want to sort your trees?"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmBroadleaves(frmNonTwoRanked.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a form allows the user to choose what kind of sorting they wish to do
Option Explicit

Private Sub cmdGoBack_Click()
    frmSort.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSort_Click()
    frmSort.Hide
    frmEverDecSort.Show
End Sub

Private Sub cmdType_Click()
    frmSort.Hide
    frmSortbytype.Show
End Sub
