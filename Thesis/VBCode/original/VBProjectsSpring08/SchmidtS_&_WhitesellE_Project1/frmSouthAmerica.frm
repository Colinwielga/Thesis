VERSION 5.00
Begin VB.Form frmSouthAmerica 
   BackColor       =   &H00C0FFFF&
   Caption         =   "South American Programs"
   ClientHeight    =   6795
   ClientLeft      =   2835
   ClientTop       =   2520
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10320
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Convert Your Money Into Chilean Pesos"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Click Here to View Projected Budget"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FFC0C0&
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
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   360
      Picture         =   "frmSouthAmerica.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3840
      ScaleHeight     =   3555
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.PictureBox picChile2 
      Height          =   3855
      Left            =   7800
      Picture         =   "frmSouthAmerica.frx":20B7
      ScaleHeight     =   3795
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00FFFFC0&
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
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblSouthAmerica 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Chile"
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
      Top             =   480
      Width           =   10335
   End
End
Attribute VB_Name = "frmSouthAmerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by Sammi and Erika
'3/13/08


Private Sub cmdInfo_Click()

picResults.Cls
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------"
picResults.Print
picResults.Print
picResults.Print "Criteria:"
picResults.Print Tab(5); "Minimum GPA of 2.5, 3 letters of"
picResults.Print Tab(5); "recommendation, and interview with"
picResults.Print Tab(5); "program director."
picResults.Print
picResults.Print "Semester:"
picResults.Print Tab(5); "You can study abroad in Chile in the fall."
picResults.Print
picResults.Print "Location:"
picResults.Print Tab(5); "Universidad Adolfo Ibanez, Viña del Mar"
picResults.Print
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------"

End Sub


Private Sub cmdBudget_Click()

MsgBox "The cost for a semester in Chile is $17,522.00 plus round-trip airfare and an estimated $1,500.00 for additional spending.", , Budget


End Sub

Private Sub cmdGoBack_Click()
frmSouthAmerica.Hide
frmPrograms.Show

End Sub


Private Sub cmdConvert_Click()
frmSouthAmerica.Hide
frmConvert.Show

End Sub
