VERSION 5.00
Begin VB.Form frmNeedleLeaves 
   BackColor       =   &H00004040&
   Caption         =   "Needles Separate or in Bundles "
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLarch 
      Caption         =   $"frmNeedleLeaves.frx":0000
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3600
      TabIndex        =   4
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton cmdPine 
      Caption         =   $"frmNeedleLeaves.frx":00C4
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturnfromNL 
      Caption         =   "To Return to First Slide:   Click Here"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdEndNL 
      Caption         =   "To End Program: Click Here"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSeparate 
      Caption         =   "If Needles Separate (As in pictures on the left side of this slide): Click Here"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Image imgPine2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4035
      Left            =   0
      Picture         =   "frmNeedleLeaves.frx":015B
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Image imgLarTam 
      BorderStyle     =   1  'Fixed Single
      Height          =   3840
      Left            =   0
      Picture         =   "frmNeedleLeaves.frx":4009D
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lblSB 
      Caption         =   "Are Needles in Bundles or Separate?"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image imgLarTam1 
      Height          =   2880
      Left            =   5880
      Picture         =   "frmNeedleLeaves.frx":1210DF
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   4635
   End
   Begin VB.Image imgPine1 
      Height          =   4380
      Left            =   6360
      Picture         =   "frmNeedleLeaves.frx":1DE221
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4620
   End
   Begin VB.Image imgSeparate2 
      Height          =   4245
      Left            =   0
      Picture         =   "frmNeedleLeaves.frx":263733
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Image imgSeparate1 
      Height          =   2985
      Left            =   0
      Picture         =   "frmNeedleLeaves.frx":29EC45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4050
   End
End
Attribute VB_Name = "frmNeedleLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmNeedleLeaves(frmBroadLeaves.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This form allows to identify if their coniferous tree is a Pine, Tamarack or Other type of tree. If it is a Tamarack or a Pine tree then it tells the scientific name as well as display a picture
Option Explicit

Private Sub cmdBundles_Click()
    frmNeedleLeaves.Hide
    frmBundles.Show
End Sub

Private Sub cmdEndNL_Click()
    End
End Sub

Private Sub cmdLarch_Click()
    imgLarTam.Visible = True
    imgPine1.Visible = False
    imgSeparate1.Visible = False
    imgSeparate2.Visible = False
    imgLarTam1.Visible = False
    cmdSeparate.Visible = False
    cmdLarch.Visible = False
    cmdPine.Visible = False
    MsgBox "Your tree is a decidudous, conifer in the genus Larix, and is commonly known as a Taramack or Larch ", , "Genus: Larix"
    imgLarTam.Visible = False
    imgPine1.Visible = True
    imgSeparate1.Visible = True
    imgSeparate2.Visible = True
    imgLarTam1.Visible = True
    cmdSeparate.Visible = True
    cmdLarch.Visible = True
    cmdPine.Visible = True
End Sub

Private Sub cmdPine_Click()
    imgPine2.Visible = True
    imgPine1.Visible = False
    imgSeparate1.Visible = False
    imgSeparate2.Visible = False
    imgLarTam1.Visible = False
    cmdSeparate.Visible = False
    cmdLarch.Visible = False
    cmdPine.Visible = False
    MsgBox "Your tree is a coniferous, evergreen in the genus Pinus, and is commonly known as a Pine ", , "Genus: Pinus"
    imgPine2.Visible = False
    imgPine1.Visible = True
    imgSeparate1.Visible = True
    imgSeparate2.Visible = True
    imgLarTam1.Visible = True
    cmdSeparate.Visible = True
    cmdLarch.Visible = True
    cmdPine.Visible = True
End Sub

Private Sub cmdReturnfromNL_Click()
    frmNeedleLeaves.Visible = False
    frmMinnesotaTrees.Visible = True
End Sub

Private Sub cmdSeparate_Click()
    frmNeedleLeaves.Hide
    frmSeparate.Show
End Sub

