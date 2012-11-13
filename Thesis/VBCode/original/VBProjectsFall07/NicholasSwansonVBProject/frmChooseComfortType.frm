VERSION 5.00
Begin VB.Form frmChooseComfortType 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Comfort Plans"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   975
      Left            =   2040
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdChooseComfort10 
      Caption         =   "Comfort 10"
      Height          =   855
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdChooseComfort12 
      Caption         =   "Comfort 12"
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdChooseComfort7 
      Caption         =   "Comfort 7"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblChooseComfortPlan 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please choose one of the comfort plans below:"
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
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "frmChooseComfortType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form gives the user a choice between the calculations pages for the
'three different Comfort Meal Plans.

Private Sub cmdBack_Click()
'Returns user to Start Page
frmChooseComfortType.Hide   'closes frmChooseComfortType
frmStartPage.Show           'opens frmStartPage
End Sub

Private Sub cmdChooseComfort10_Click()
'Advances user to the Comfort 10 Plan calculations page
frmChooseComfortType.Hide       'closes frmChooseComfortType
frmComfort10.Show               'opens frmComfort10
End Sub

Private Sub cmdChooseComfort12_Click()
'Advances user to the Comfort 12 Plan calculations page
frmChooseComfortType.Hide       'closes frmChooseComfortType
frmComfort12.Show               'opens frmComfort12
End Sub

Private Sub cmdChooseComfort7_Click()
'Advances user to the Comfort 7 Plan calculations page
frmChooseComfortType.Hide       'closes frmChooseComfortType
frmComfort7.Show                'open frmComfort7
End Sub

Private Sub Form_Load()
'loads SJU logo
picResults.Picture = LoadPicture(App.Path & "\johnnieslogo.gif")
'centers form on computer screen upon loading
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
