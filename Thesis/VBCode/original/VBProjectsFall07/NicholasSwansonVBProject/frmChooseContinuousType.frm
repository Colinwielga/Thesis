VERSION 5.00
Begin VB.Form frmChooseContinuousType 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Continuous Plans"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   1560
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdChooseMonThruFri 
      Caption         =   "Monday Thru Friday"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdChoose7DaysPerWeek 
      Caption         =   "7 Days per Week"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblChooseContinuousPlan 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please choose one of the continuous plans below:"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmChooseContinuousType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form gives the user a choice between the Continuous Plan 7 Days a week and
'the Continuous Plan Monday thru Friday calculations pages.

Private Sub cmdBack_Click()
'Returns user to Start Page
frmChooseContinuousType.Hide    'closes frmChooseContinuousType
frmStartPage.Show               'opens frmStartPage
End Sub

Private Sub cmdChoose7DaysPerWeek_Click()
'Advances user to the Continuous Plan 7 Days a Week calculations page
frmChooseContinuousType.Hide    'closes frmChooseContinuousType
frm7DaysAWeek.Show              'opens frm7DaysAWeek
End Sub

Private Sub cmdChooseMonThruFri_Click()
'Advances user to the Continuous Plan Monday thru Friday calculations page
frmChooseContinuousType.Hide    'closes frmChooseContinuousType
frmMonThruFri.Show              'opens frmMonThruFri
End Sub

Private Sub Form_Load()
'loads SJU logo
picResults.Picture = LoadPicture(App.Path & "\johnnieslogo.gif")
'centers form on computer screen upon loading
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
