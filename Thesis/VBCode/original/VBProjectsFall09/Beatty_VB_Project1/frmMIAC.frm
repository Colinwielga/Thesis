VERSION 5.00
Begin VB.Form frmMIAC 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmMIAC.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCitations 
      Caption         =   "Click for Citations"
      Height          =   660
      Left            =   3120
      TabIndex        =   19
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuestions 
      Caption         =   "Click here for some Questions"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Click to see General Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmdUST 
      Caption         =   "University of St. Thomas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOlaf 
      Caption         =   "St. Olaf College"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdMary 
      Caption         =   "St. Mary's University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSJU 
      Caption         =   "St. John's University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCate 
      Caption         =   "St. Catherine University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCSB 
      Caption         =   "College of St. Benedict"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdMac 
      Caption         =   "Macalaster College"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdHamline 
      Caption         =   "Hamline University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdGust 
      Caption         =   "Gustavus Adolphus College"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdConcordia 
      Caption         =   "Concordia College"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCarleton 
      Caption         =   "Carleton College"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBethel 
      Caption         =   "Bethel University"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Picture         =   "frmMIAC.frx":1D6A
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton cmdAugsburg 
      BackColor       =   &H8000000D&
      Caption         =   "Augsburg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Click on a College to learn a bit more  ==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblnames 
      BackColor       =   &H8000000D&
      Caption         =   "Teams:"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackColor       =   &H8000000D&
      Caption         =   "Welcome to the M.I.A.C. (Minnesota Intercollegiate Athletic Conference)"
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmMIAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Project name: quick Facts about the MIAC'
    'Form name:MIAC
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about schools around the MIAC'
    
Option Explicit

Private Sub cmdAugsburg_Click()
    frmAugsburg.Show
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdBethel_Click()
    frmAugsburg.Hide
    frmBethel.Show
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdCarleton_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Show
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdCate_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Show
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdCitations_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
    frmCitations.Show
End Sub

Private Sub cmdConcordia_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Show
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdCSB_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Show
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdGust_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Show
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdHamline_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Show
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdInfo_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Show
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide

End Sub

Private Sub cmdMac_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Show
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdMary_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Show
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdOlaf_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Show
    frmUST.Hide
    frmQuestions.Hide
End Sub

Private Sub cmdQuestions_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSJU_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Show
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Hide
    frmQuestions.Hide
    
End Sub

Private Sub cmdUST_Click()
    frmAugsburg.Hide
    frmBethel.Hide
    frmCarleton.Hide
    frmConcordia.Hide
    frmCSB.Hide
    frmGustavus.Hide
    frmHamline.Hide
    frmInfo.Hide
    frmMacalaster.Hide
    frmSJU.Hide
    frmStCates.Hide
    frmStMary.Hide
    frmStOLaf.Hide
    frmUST.Show
End Sub
