VERSION 5.00
Begin VB.Form frmSimple 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Leaves Simple and Lobed "
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndSL 
      Caption         =   "To End Program: Click Here "
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   1
      Top             =   4680
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturnfromSL 
      Caption         =   "If none of these: Click Here to return to beginning of program"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblSimpLobe 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Simple and Lobed Leaves "
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   3360
      Width           =   7095
   End
   Begin VB.Image imgMul2 
      Height          =   4335
      Left            =   0
      Picture         =   "frmSimple.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.Image imgEurope 
      Height          =   5100
      Left            =   -240
      Picture         =   "frmSimple.frx":5B062
      Top             =   120
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Image imgOak2 
      Height          =   5625
      Left            =   240
      Picture         =   "frmSimple.frx":9BC94
      Top             =   -240
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Label lblMaples 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Leaves palmately lobed and palmately veined opposite leaves (Maples): Click Picture Above for more Info"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7800
      TabIndex        =   5
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label lblMul 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSimple.frx":E0A66
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label lblOaks 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laterally lobed and Palmately veined leaves (Oak): Click Picture below for more info"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image imgMaple2 
      Height          =   3990
      Left            =   480
      Picture         =   "frmSimple.frx":E0B09
      Top             =   1440
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblBirchSL 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Braches drooping, bark white (European Brich): Click Picture Below for more info"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image imgMaple 
      Height          =   3300
      Left            =   5760
      Picture         =   "frmSimple.frx":105DA3
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   3945
   End
   Begin VB.Image imgMulberry 
      Height          =   2895
      Left            =   960
      Picture         =   "frmSimple.frx":18AC25
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   4020
   End
   Begin VB.Image imgOak 
      Height          =   2625
      Left            =   5520
      Picture         =   "frmSimple.frx":2815A3
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image imgBirchSL 
      Height          =   2715
      Left            =   720
      Picture         =   "frmSimple.frx":2AF445
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "frmSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmBroadleaves(frmNonTwoRanked.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a form allows the final the user to categorize the different types of simple leaves
Option Explicit
Private Sub cmdEndSL_Click()
    End
End Sub

Private Sub cmdReturnfromSL_Click()
    frmSimple.Hide
    frmMinnesotaTrees.Show
End Sub


Private Sub imgBirchSL_Click()
    imgEurope.Visible = True
    imgMaple.Visible = False
    imgMulberry.Visible = False
    lblOaks.Visible = False
    lblMaples.Visible = False
    lblMul.Visible = False
    lblBirchSL.Visible = False
    lblSimpLobe.Visible = False
    imgMulberry.Visible = False
    imgBirchSL.Visible = False
    imgOak.Visible = False
    cmdReturnfromSL.Visible = False
    cmdEndSL.Visible = False
    MsgBox " Your tree is a deciduous tree in the genus Betula, and is commonly known as the European Birch ", , "Genus: Betula"
    imgEurope.Visible = False
    imgMaple.Visible = True
    imgMulberry.Visible = True
    lblOaks.Visible = True
    lblMaples.Visible = True
    lblMul.Visible = True
    lblBirchSL.Visible = True
    lblSimpLobe.Visible = True
    imgMulberry.Visible = True
    imgBirchSL.Visible = True
    imgOak.Visible = True
    cmdReturnfromSL.Visible = True
    cmdEndSL.Visible = True
End Sub
Private Sub imgMaple_Click()
    imgMaple2.Visible = True
    imgMaple.Visible = False
    imgMulberry.Visible = False
    lblOaks.Visible = False
    lblMaples.Visible = False
    lblMul.Visible = False
    lblBirchSL.Visible = False
    lblSimpLobe.Visible = False
    imgMulberry.Visible = False
    imgBirchSL.Visible = False
    imgOak.Visible = False
    cmdReturnfromSL.Visible = False
    cmdEndSL.Visible = False
    MsgBox " Your tree is a deciduous tree in the genus Acer, and is commonly known as a Maple ", , "Genus: Acer"
    imgMaple2.Visible = False
    imgMaple.Visible = True
    imgMulberry.Visible = True
    lblOaks.Visible = True
    lblMaples.Visible = True
    lblMul.Visible = True
    lblBirchSL.Visible = True
    lblSimpLobe.Visible = True
    imgMulberry.Visible = True
    imgBirchSL.Visible = True
    imgOak.Visible = True
    cmdReturnfromSL.Visible = True
    cmdEndSL.Visible = True
End Sub

Private Sub imgMulberry_Click()
    imgMul2.Visible = True
    imgMaple.Visible = False
    imgMulberry.Visible = False
    lblOaks.Visible = False
    lblMaples.Visible = False
    lblMul.Visible = False
    lblBirchSL.Visible = False
    lblSimpLobe.Visible = False
    imgMulberry.Visible = False
    imgBirchSL.Visible = False
    imgOak.Visible = False
    cmdReturnfromSL.Visible = False
    cmdEndSL.Visible = False
    MsgBox " Your tree is a deciduous tree in the genus Morus, and is commonly known as a Mulberry", , "Genus: Morus"
    imgMul2.Visible = False
    imgMaple.Visible = True
    imgMulberry.Visible = True
    lblOaks.Visible = True
    lblMaples.Visible = True
    lblMul.Visible = True
    lblBirchSL.Visible = True
    lblSimpLobe.Visible = True
    imgMulberry.Visible = True
    imgBirchSL.Visible = True
    imgOak.Visible = True
    cmdReturnfromSL.Visible = True
    cmdEndSL.Visible = True
End Sub

Private Sub imgOak_Click()
    imgOak2.Visible = True
    imgMaple.Visible = False
    imgMulberry.Visible = False
    lblOaks.Visible = False
    lblMaples.Visible = False
    lblMul.Visible = False
    lblBirchSL.Visible = False
    lblSimpLobe.Visible = False
    imgMulberry.Visible = False
    imgBirchSL.Visible = False
    imgOak.Visible = False
    cmdReturnfromSL.Visible = False
    cmdEndSL.Visible = False
    MsgBox " Your tree is a deciduous tree in the genus Quercus, and is commonly known as a Oak", , "Genus: Quercus"
    imgOak2.Visible = False
    imgMaple.Visible = True
    imgMulberry.Visible = True
    lblOaks.Visible = True
    lblMaples.Visible = True
    lblMul.Visible = True
    lblBirchSL.Visible = True
    lblSimpLobe.Visible = True
    imgMulberry.Visible = True
    imgBirchSL.Visible = True
    imgOak.Visible = True
    cmdReturnfromSL.Visible = True
    cmdEndSL.Visible = True
End Sub
