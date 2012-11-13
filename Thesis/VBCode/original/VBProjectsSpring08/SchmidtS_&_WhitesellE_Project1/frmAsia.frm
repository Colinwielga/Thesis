VERSION 5.00
Begin VB.Form frmAsia 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Asian Programs"
   ClientHeight    =   6810
   ClientLeft      =   2730
   ClientTop       =   2415
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10380
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Covert Your Money"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   960
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H0080C0FF&
      Caption         =   "View Program Budgets"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Display Program Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H0080C0FF&
      Caption         =   "Display Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox picAsia 
      Height          =   3135
      Left            =   5640
      Picture         =   "frmAsia.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asian Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10455
   End
End
Attribute VB_Name = "frmAsia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This form provides information on the Chinese and Japan programs.
'Written 3/25/08 by Sammi




Private Sub cmdGoBack_Click()
frmAsia.Hide
frmPrograms.Show

End Sub

Private Sub cmdList_Click()
picResults.Cls
picResults.Print "------------------------------------------------------------------------------"
picResults.Print
picResults.Print Tab(20); "China"
picResults.Print Tab(20); "Japan"
picResults.Print
picResults.Print "------------------------------------------------------------------------------"

End Sub

Private Sub cmdInfo_Click()
'displays program details in the picture box


picResults.Cls
picResults.Print " -----------------------------------------------------------------------------"

picResults.Print "Criteria:"
picResults.Print
picResults.Print Tab(4); "Minimum GPA of 2.5, 3 letters of"
picResults.Print Tab(4); " recommendation, interview with"
picResults.Print Tab(4); " program director."
picResults.Print
picResults.Print "Available Semesters:"
picResults.Print
picResults.Print Tab(4); "Both programs are only available in the fall."
picResults.Print
picResults.Print " -----------------------------------------------------------------------------"


End Sub


Private Sub cmdBudget_Click()

'displays cost of programs in a message box

MsgBox "The cost for a semester in China or Japan is $17,522.00, plus round-trip airfare and an estimated $2,000.00 for additional spending.", , "Budget"



End Sub


Private Sub cmdConvert_Click()
frmAsia.Hide
frmConvert.Show
End Sub
