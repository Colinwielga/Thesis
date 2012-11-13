VERSION 5.00
Begin VB.Form frmNonTwoRanked 
   BackColor       =   &H0000C0C0&
   Caption         =   "Trees with leaves around the twig"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EndNon2 
      Caption         =   "To End Program Click Here"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8640
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturnNon2 
      Caption         =   "If you want to Return to the beginning of the program: Click Here"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Image imgPears2 
      Height          =   4065
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.Image imgApple2 
      Height          =   4680
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":89586
      Stretch         =   -1  'True
      Top             =   -360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image imgWillow2 
      Height          =   4680
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":C9FC8
      Stretch         =   -1  'True
      Top             =   -120
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.Image imgCherry2 
      Height          =   4200
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":12520A
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Label lblCrabApples 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buds not flattened against twig, leaves hairy beneath, broadly ovate (apples, crapapples)"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6240
      TabIndex        =   4
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label lblCherry 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buds not flattened against twig, leaves smooth, finely and evenly toothed, broadly lancelolate to ovate (Cherries)"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   6240
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label lblPears 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buds not flattened against twig, leaves hairy beneath, broadly ovate (pears)"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4455
   End
   Begin VB.Label lblWillow 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buds flattened against twigs; leaves lanceolate (Willow):  Click Here for more info"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label imgTitleNon2 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Toothed Leaves with leaves emerging from many directions from the twig"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image imgPear 
      Height          =   3600
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":19364C
      Top             =   3000
      Width           =   4800
   End
   Begin VB.Image imgCrabApples 
      Height          =   3945
      Left            =   6000
      Picture         =   "frmNonTwoRanked.frx":1CBA8E
      Top             =   2880
      Width           =   5250
   End
   Begin VB.Image imgCherry 
      Height          =   3000
      Left            =   6000
      Picture         =   "frmNonTwoRanked.frx":20F394
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4770
   End
   Begin VB.Image imgWillow 
      Height          =   3255
      Left            =   0
      Picture         =   "frmNonTwoRanked.frx":36E376
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   4785
   End
End
Attribute VB_Name = "frmNonTwoRanked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmNonTwoRanked(frmNonTwoRanked.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a form allows the final identification of some popular fruit trees
Option Explicit
Private Sub cmdReturnNon2_Click()
    frmNonTwoRanked.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub EndNon2_Click()
    End
End Sub
Private Sub lblCherry_Click()
    imgCherry2.Visible = True
    lblCherry.Visible = False
    lblWillow.Visible = False
    lblPears.Visible = False
    imgCherry.Visible = False
    imgWillow.Visible = False
    imgPear.Visible = False
    lblCrabApples.Visible = False
    imgCrabApples.Visible = False
    imgTitleNon2.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Prunus, and is commonly known as a Cherry Tree ", , "Genus: Prunus"
    imgCherry2.Visible = False
    lblCherry.Visible = True
    lblWillow.Visible = True
    lblPears.Visible = True
    imgCherry.Visible = True
    imgWillow.Visible = True
    imgPear.Visible = True
    lblCrabApples.Visible = True
    imgCrabApples.Visible = True
    imgTitleNon2.Visible = True
End Sub

Private Sub lblCrabApples_Click()
    imgApple2.Visible = True
    lblCherry.Visible = False
    lblWillow.Visible = False
    lblPears.Visible = False
    imgCherry.Visible = False
    imgWillow.Visible = False
    imgPear.Visible = False
    lblCrabApples.Visible = False
    imgCrabApples.Visible = False
    imgTitleNon2.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Malus, and is commonly known as an Apple or Crabapple Tree ", , "Genus: Malus"
    imgApple2.Visible = False
    lblCherry.Visible = True
    lblWillow.Visible = True
    lblPears.Visible = True
    imgCherry.Visible = True
    imgWillow.Visible = True
    imgPear.Visible = True
    lblCrabApples.Visible = True
    imgCrabApples.Visible = True
    imgTitleNon2.Visible = True
End Sub

Private Sub lblPears_Click()
    imgPears2.Visible = True
    lblCherry.Visible = False
    lblWillow.Visible = False
    lblPears.Visible = False
    imgCherry.Visible = False
    imgWillow.Visible = False
    imgPear.Visible = False
    lblCrabApples.Visible = False
    imgCrabApples.Visible = False
    imgTitleNon2.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Pyrus, and is commonly known as a Pear Tree ", , "Genus: Pyrus"
    imgPears2.Visible = False
    lblCherry.Visible = True
    lblWillow.Visible = True
    lblPears.Visible = True
    imgCherry.Visible = True
    imgWillow.Visible = True
    imgPear.Visible = True
    lblCrabApples.Visible = True
    imgCrabApples.Visible = True
    imgTitleNon2.Visible = True
End Sub

Private Sub lblWillow_Click()
    imgWillow2.Visible = True
    lblCherry.Visible = False
    lblWillow.Visible = False
    lblPears.Visible = False
    imgCherry.Visible = False
    imgWillow.Visible = False
    imgPear.Visible = False
    lblCrabApples.Visible = False
    imgCrabApples.Visible = False
    imgTitleNon2.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Salix, and is commonly known as a Willow ", , "Genus: Salix"
    imgWillow2.Visible = False
    lblCherry.Visible = True
    lblWillow.Visible = True
    lblPears.Visible = True
    imgCherry.Visible = True
    imgWillow.Visible = True
    imgPear.Visible = True
    lblCrabApples.Visible = True
    imgCrabApples.Visible = True
    imgTitleNon2.Visible = True
End Sub
