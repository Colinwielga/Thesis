VERSION 5.00
Begin VB.Form frmChoosing 
   Caption         =   "Choose Your Pokemon"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10245
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   435
      Left            =   13080
      TabIndex        =   6
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdVenusaur 
      Cancel          =   -1  'True
      Caption         =   "Venusaur"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   9360
      Width           =   2895
   End
   Begin VB.CommandButton cmdBlastoise 
      Caption         =   "Blastoise"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdCharizard 
      Caption         =   "Charizard "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   4440
      Width           =   2895
   End
   Begin VB.PictureBox picVenusaur 
      Height          =   3495
      Left            =   5640
      Picture         =   "Form1.frx":6225
      ScaleHeight     =   3435
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   5760
      Width           =   3855
   End
   Begin VB.PictureBox picBlastoise 
      Height          =   3495
      Left            =   9960
      Picture         =   "Form1.frx":ACC0
      ScaleHeight     =   3435
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.PictureBox picCharizard 
      Height          =   3375
      Left            =   840
      Picture         =   "Form1.frx":DAE5
      ScaleHeight     =   3315
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblGrass 
      BackStyle       =   0  'Transparent
      Caption         =   "Grass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblWater 
      BackStyle       =   0  'Transparent
      Caption         =   "Water"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   11160
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblFire 
      BackStyle       =   0  'Transparent
      Caption         =   "Fire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblWarning2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remember:"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   6360
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "      Fire > Grass        Grass > Water       Water > Fire"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   5880
      TabIndex        =   11
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label lblPeriod 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblName1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblName2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblChoose 
      BackStyle       =   0  'Transparent
      Caption         =   "    Choose            Your         Pokémon!"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   5880
      TabIndex        =   7
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "frmChoosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pokemon Project
'frmChoosing
'Eugene Zheng
'10/11/2009
'This is where the user decides which pokemon he/she wishes to use
'There are three different results: Charizard, Blastoise, or Venusaur

'Battles are governed by each pokemon's individual stats.
'For example: Charizard is faster than Venusaur, thus I wrote the code to enable Charizard to attack first


Option Explicit


Private Sub cmdBlastoise_Click()
'This button chooses Blastoise
frmChoosing.Hide
frmBlastoiseChosen.Show
End Sub

Private Sub cmdCharizard_Click()
'This button chooses Charizard
frmChoosing.Hide
frmCharizardChosen.Show
End Sub

Private Sub cmdQuit_Click()
'simple quit button
End
End Sub

Private Sub cmdVenusaur_Click()
'This button chooses Venusaur
frmChoosing.Hide
frmVenusaurChosen.Show
End Sub

Private Sub Form_Load()
Dim UserFirstName As String
Dim UserLastName As String

'Before the everything is seen, we want to find out the user names
UserFirstName = InputBox("Enter Your First Name", "Name")
UserLastName = InputBox("Enter Your Last Name", "Name")

'We only use the first initial for simplicity sake
If Len(UserFirstName) > 0 Then
    lblPeriod.Caption = "."
End If

'Use the labels to print the names
    lblName1.Caption = Left(UserFirstName, 1)
    lblName2.Caption = UserLastName

lblFire.ForeColor = vbRed
lblWater.ForeColor = vbBlue
lblGrass.ForeColor = vbGreen
End Sub


