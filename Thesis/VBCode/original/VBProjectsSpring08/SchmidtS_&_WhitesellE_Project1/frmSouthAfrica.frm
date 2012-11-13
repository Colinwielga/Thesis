VERSION 5.00
Begin VB.Form frmSouthAfrica 
   BackColor       =   &H00FFFFC0&
   Caption         =   "SouthAfrica"
   ClientHeight    =   6825
   ClientLeft      =   2730
   ClientTop       =   2415
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10350
   Begin VB.PictureBox picSAfrica2 
      Height          =   1575
      Left            =   720
      Picture         =   "frmSouthAfrica.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   4080
      Width           =   2415
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
      Height          =   735
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Convert Your Money into South African Rand"
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
      TabIndex        =   5
      Top             =   5880
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3720
      ScaleHeight     =   3555
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click Here to See Projected Budget"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.PictureBox picSAfrica 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   360
      Picture         =   "frmSouthAfrica.frx":1EC7
      ScaleHeight     =   1815
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblSouthAfrica 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "South Africa"
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
      Top             =   720
      Width           =   10335
   End
End
Attribute VB_Name = "frmSouthAfrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by Sammi and Erika
'3/13/03


Private Sub cmdInfo_Click()

picResults.Cls
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------"
picResults.Print
picResults.Print "Criteria:"
picResults.Print Tab(5); "Minimum GPA of 2.5, 3 letters of recommendation,"
picResults.Print Tab(5); "interview with program director"
picResults.Print
picResults.Print "Semester:"
picResults.Print Tab(5); "You can study abroad in South Africa in the spring."
picResults.Print
picResults.Print "Location:"
picResults.Print Tab(5); "Nelson Mandela Metropolitan University,"
picResults.Print Tab(5); "Port Elizabeth"
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------"

End Sub


Private Sub cmdBudget_Click()

MsgBox "The cost for a semester in South Africa is $17,837.00 plus round-trip airfare and $2,000.00 for additional spending.", , Budget



End Sub

Private Sub cmdGoBack_Click()
frmSouthAfrica.Hide
frmPrograms.Show

End Sub

Private Sub cmdConvert_Click()
frmSouthAfrica.Hide
frmConvert.Show

End Sub
