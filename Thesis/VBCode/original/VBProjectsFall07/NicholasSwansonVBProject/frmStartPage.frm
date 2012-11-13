VERSION 5.00
Begin VB.Form frmStartPage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Meal Plan Projection Calculator"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   3600
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdSources 
      Caption         =   "Sources"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Food Fight 2007"
      Height          =   735
      Left            =   8040
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdBlockPlan 
      Caption         =   "Block Plan"
      Height          =   1095
      Left            =   5160
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdContinuousPlans 
      Caption         =   "Continuous Plans"
      Height          =   1095
      Left            =   2880
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdApartmentPlan 
      Caption         =   "Apartment Plan"
      Height          =   1095
      Left            =   7440
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdComfortPlans 
      Caption         =   "Comfort Plans"
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblFight 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fight!"
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   6480
      TabIndex        =   9
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblFood 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmStartPage.frx":0000
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   1560
      TabIndex        =   0
      Top             =   2760
      Width           =   6735
   End
End
Attribute VB_Name = "frmStartPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form greets the user as the Start Page.  From here, various buttons branch the user
'in several different directions, depending on for which meal plans s/he would like to make
'calculations.  It also contains a command button linking the user to a brief list of code
'and images sources, as well as having the only quite button in the program.  This encourages
'backtracking and reevaluating of results, as well as calculations for other plans, should the
'user be trying to make a decision between them.

Private Sub cmdApartmentPlan_Click()
'Advances user to Apartment Plan calculations page
frmStartPage.Hide       'closes frmStartPage
frmApartmentPlan.Show   'opens frmApartmentPlan
End Sub

Private Sub cmdBlockPlan_Click()
'Advances user to Block Plan calculations page
frmStartPage.Hide       'closes frmStartPage
frmBlockPlan.Show       'opens frmBlockPlan
End Sub

Private Sub cmdComfortPlans_Click()
'Advances user to a page where s/he chooses for which comfort plan to do calculations
frmStartPage.Hide           'closes frmStartPage
frmChooseComfortType.Show   'opens frmChooseComfortType
End Sub

Private Sub cmdContinuousPlans_Click()
'Advances user to a page where s/he chooses for which continuous plan to do calculations
frmStartPage.Hide               'closes frmStartPage
frmChooseContinuousType.Show    'opens frmChooseContinuousType
End Sub

Private Sub cmdQuit_Click()
'Closes the program
End
End Sub

Private Sub cmdSources_Click()
'Advances user to a page of code and image sources
frmSources.Show     'opens frmSources
End Sub

Private Sub Form_Load()
'loads SJU logo
picResults.Picture = LoadPicture(App.Path & "\logo-sju.gif")
'centers form on computer screen upon loading
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
