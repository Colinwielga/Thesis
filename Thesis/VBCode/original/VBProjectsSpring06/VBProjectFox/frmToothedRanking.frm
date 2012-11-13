VERSION 5.00
Begin VB.Form frmToothedRanked 
   BackColor       =   &H00000000&
   Caption         =   "Toothed Simple Leaves"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdEndTR 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   4
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturnTR 
      Caption         =   "If none of these: Click here"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Image imgHack 
      Height          =   1800
      Left            =   1800
      Picture         =   "frmToothedRanking.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Image imgHackberry2 
      Height          =   3720
      Left            =   -360
      Picture         =   "frmToothedRanking.frx":5B242
      Stretch         =   -1  'True
      Top             =   -240
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Image imgChina2 
      Height          =   3780
      Left            =   -120
      Picture         =   "frmToothedRanking.frx":13C284
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgIron2 
      Height          =   3225
      Left            =   480
      Picture         =   "frmToothedRanking.frx":198C06
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Image imgGotoNon2 
      Height          =   1515
      Left            =   4200
      Picture         =   "frmToothedRanking.frx":2D23E8
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2520
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   4200
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label lblvarious 
      BackStyle       =   0  'Transparent
      Caption         =   "If leaves emerging from many directions click picture below"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3840
      TabIndex        =   10
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Image imgIron 
      Height          =   2160
      Left            =   7800
      Picture         =   "frmToothedRanking.frx":3979AA
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2160
   End
   Begin VB.Label lblIron 
      BackColor       =   &H80000012&
      Caption         =   $"frmToothedRanking.frx":4117EC
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   5640
      TabIndex        =   9
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Label LblChina 
      BackColor       =   &H80000012&
      Caption         =   $"frmToothedRanking.frx":4118A0
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Image imgChina 
      Height          =   2280
      Left            =   2280
      Picture         =   "frmToothedRanking.frx":411930
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image imgElm2 
      Height          =   3660
      Left            =   0
      Picture         =   "frmToothedRanking.frx":846EF2
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1560
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lbl2rankednolop 
      BackStyle       =   0  'Transparent
      Caption         =   "Leaves two ranked and  slightly or not at all lopsided at the base"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   9495
   End
   Begin VB.Image imgElm 
      Height          =   1695
      Left            =   7440
      Picture         =   "frmToothedRanking.frx":927F34
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2505
   End
   Begin VB.Label lblHack 
      BackStyle       =   0  'Transparent
      Caption         =   "Leaves very thin palmately three veined (Hackberry): Click picture below for more information"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label lblElm 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Leaves thick, often rough on upper surfaces (Elms): Click picture below for more information"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6000
      TabIndex        =   5
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label lblTitle2TS 
      BackColor       =   &H80000012&
      Caption         =   "Leaves emerging from various directions from the twig"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   6600
      Width           =   8535
   End
   Begin VB.Label lbl2rankedlop 
      BackColor       =   &H80000012&
      Caption         =   "Leaves apparently two-ranked and lopsided at base"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label lblTitleTSL 
      BackStyle       =   0  'Transparent
      Caption         =   "Toothed Simple Leaves "
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmToothedRanked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmToothedRanked(frmToothedRanked.frm)
'Author: Kelly Fox
'Date Written:3/22/2006
'This is a form allows the final identification of some toothed simple leaves
Option Explicit

Private Sub cmdEndTR_Click()
    End
End Sub


Private Sub cmdReturnTR_Click()
    frmToothedRanked.Hide
    frmMinnesotaTrees.Show
End Sub
Private Sub imgChina_Click()
    imgChina2.Visible = True
    imgElm.Visible = False
    imgIron.Visible = False
    imgHack.Visible = False
    imgChina.Visible = False
    imgGotoNon2.Visible = False
    lblElm.Visible = False
    lblIron.Visible = False
    lblHack.Visible = False
    LblChina.Visible = False
    lblTitle2TS.Visible = False
    lblvarious.Visible = False
    lbl2rankedlop.Visible = False
    lbl2rankednolop.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Ulmus, and is commonly known as a Chinese Elm", , "Genus: Ulmus"
    imgChina2.Visible = False
    imgElm.Visible = True
    imgIron.Visible = True
    imgHack.Visible = True
    imgChina.Visible = True
    imgGotoNon2.Visible = True
    lblElm.Visible = True
    lblIron.Visible = True
    lblHack.Visible = True
    LblChina.Visible = True
    lblTitle2TS.Visible = True
    lblvarious.Visible = True
    lbl2rankedlop.Visible = True
    lbl2rankednolop.Visible = True
End Sub

Private Sub imgElm_Click()
    imgElm2.Visible = True
    imgElm.Visible = False
    imgIron.Visible = False
    imgHack.Visible = False
    imgChina.Visible = False
    imgGotoNon2.Visible = False
    lblElm.Visible = False
    lblIron.Visible = False
    lblHack.Visible = False
    LblChina.Visible = False
    lblTitle2TS.Visible = False
    lblvarious.Visible = False
    lbl2rankedlop.Visible = False
    lbl2rankednolop.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Ulmus, and is commonly known as a Elm ", , "Genus: Ulmus"
    imgElm2.Visible = False
    imgElm.Visible = True
    imgIron.Visible = True
    imgHack.Visible = True
    imgChina.Visible = True
    imgGotoNon2.Visible = True
    lblElm.Visible = True
    lblIron.Visible = True
    lblHack.Visible = True
    LblChina.Visible = True
    lblTitle2TS.Visible = True
    lblvarious.Visible = True
    lbl2rankedlop.Visible = True
    lbl2rankednolop.Visible = True
End Sub

Private Sub imgGotoNon2_Click()
    frmToothedRanked.Hide
    frmNonTwoRanked.Show
End Sub

   
Private Sub imgHack_Click()
    imgHackberry2.Visible = True
    imgElm.Visible = False
    imgIron.Visible = False
    imgHack.Visible = False
    imgChina.Visible = False
    imgGotoNon2.Visible = False
    lblElm.Visible = False
    lblIron.Visible = False
    lblHack.Visible = False
    LblChina.Visible = False
    lblTitle2TS.Visible = False
    lblvarious.Visible = False
    lbl2rankedlop.Visible = False
    lbl2rankednolop.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Celtis, and is commonly known as a Hackberry", , "Genus: Celtis"
    imgHackberry2.Visible = False
    imgElm.Visible = True
    imgIron.Visible = True
    imgHack.Visible = True
    imgChina.Visible = True
    imgGotoNon2.Visible = True
    lblElm.Visible = True
    lblIron.Visible = True
    lblHack.Visible = True
    LblChina.Visible = True
    lblTitle2TS.Visible = True
    lblvarious.Visible = True
    lbl2rankedlop.Visible = True
    lbl2rankednolop.Visible = True
End Sub

Private Sub imgIron_Click()
    imgIron2.Visible = True
    imgElm.Visible = False
    imgIron.Visible = False
    imgHack.Visible = False
    imgChina.Visible = False
    imgGotoNon2.Visible = False
    lblElm.Visible = False
    lblIron.Visible = False
    lblHack.Visible = False
    LblChina.Visible = False
    lblTitle2TS.Visible = False
    lblvarious.Visible = False
    lbl2rankedlop.Visible = False
    lbl2rankednolop.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Ostrya, and is commonly known as a Ironwood or Hop hornbeam ", , "Genus: Ostrya"
    imgIron2.Visible = False
    imgElm.Visible = True
    imgIron.Visible = True
    imgHack.Visible = True
    imgChina.Visible = True
    imgGotoNon2.Visible = True
    lblElm.Visible = True
    lblIron.Visible = True
    lblHack.Visible = True
    LblChina.Visible = True
    lblTitle2TS.Visible = True
    lblvarious.Visible = True
    lbl2rankedlop.Visible = True
    lbl2rankednolop.Visible = True
End Sub

