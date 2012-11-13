VERSION 5.00
Begin VB.Form frmAustralia 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Australia"
   ClientHeight    =   6270
   ClientLeft      =   3030
   ClientTop       =   2715
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9915
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Convert Your Money into Australian Dollars"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to View  Projected Budjet "
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.PictureBox picAustralia 
      Height          =   1695
      Left            =   1080
      Picture         =   "frmAustralia.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   480
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
      Height          =   3375
      Left            =   3600
      ScaleHeight     =   3315
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Display Program Details"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0C0FF&
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblAustralia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "                 Australia"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9975
   End
End
Attribute VB_Name = "frmAustralia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'This form shows information on the Australian program.
'Written 3/12/08 by Erika

Private Sub cmdInfo_Click()

picResults.Cls
picResults.Print "----------------------------------------------------------------------------------------------"
picResults.Print "Criteria:"
picResults.Print Tab(5); "Minimum GPA of 2.5, 3 letters of recommendation,"
picResults.Print Tab(5); "interview with program director"
picResults.Print
picResults.Print "Semester:"
picResults.Print Tab(5); "You can study abroad in Austrialia in the fall,"
picResults.Print Tab(5); "AND in the spring."
picResults.Print
picResults.Print "Location:"
picResults.Print Tab(5); "University of Notre Dame Australia, Fremantle"
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------"

End Sub


Private Sub cmdBudget_Click()

MsgBox "The cost for a semester in Australia is $18,672.00 plus round-trip airfare and $2,000.00 for additional spending.", , Budget



End Sub

Private Sub cmdGoBack_Click()
frmAustralia.Hide
frmPrograms.Show

End Sub


Private Sub cmdConvert_Click()
frmAustralia.Hide
frmConvert.Show
End Sub
