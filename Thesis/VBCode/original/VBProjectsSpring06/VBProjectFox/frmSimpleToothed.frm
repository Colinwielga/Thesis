VERSION 5.00
Begin VB.Form frmSimpleToothed 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identification of Toothed Leaves"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndST 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturnST 
      Caption         =   "Return to First Slide "
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Image imgPoplars 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSimpleToothed.frx":0000
      Top             =   -240
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image imgBass2 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSimpleToothed.frx":29E92
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image imgPlums2 
      Height          =   3360
      Left            =   -360
      Picture         =   "frmSimpleToothed.frx":5EA94
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Image imgBirch2 
      Height          =   4080
      Left            =   0
      Picture         =   "frmSimpleToothed.frx":13FAD6
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Image imgBirch 
      BorderStyle     =   1  'Fixed Single
      Height          =   2085
      Left            =   6000
      Picture         =   "frmSimpleToothed.frx":1EF798
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1770
   End
   Begin VB.Image imgHaw2 
      Height          =   3750
      Left            =   0
      Picture         =   "frmSimpleToothed.frx":21702E
      Top             =   -120
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image imgPlums 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   2880
      Picture         =   "frmSimpleToothed.frx":244ED0
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   2160
   End
   Begin VB.Label lblPlums 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSimpleToothed.frx":2B5712
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   2640
      TabIndex        =   12
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblHaw 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Has sharp thorns and broad leaves (Hawthorne):Click below for more information"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Image imgHawthorne 
      BorderStyle     =   1  'Fixed Single
      Height          =   1545
      Left            =   120
      Picture         =   "frmSimpleToothed.frx":2B57A1
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image imgMoreToothedLeaves 
      BorderStyle     =   1  'Fixed Single
      Height          =   2160
      Left            =   8760
      Picture         =   "frmSimpleToothed.frx":2EC02B
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1560
   End
   Begin VB.Label lblPetmore 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Many petioles of 1/12 inches or more in length"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label lblPoplar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSimpleToothed.frx":34726D
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   4800
      TabIndex        =   9
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Image imgPoplar 
      BorderStyle     =   1  'Fixed Single
      Height          =   1845
      Left            =   8520
      Picture         =   "frmSimpleToothed.frx":347300
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Toothed Leaves"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblBass 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Leaves lopsided at base, heart-shaped, and palmately veined (Basswoods/Linders): Click the picture above for more information"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Image imgBass 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   240
      Picture         =   "frmSimpleToothed.frx":35EA0A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label lblNotPaper 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bark Not Peeling in Papery Layers: Click Below"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   8160
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label lblBirch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bark Peeling in Papery Layers (Birches) Click Below for more information"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   5640
      TabIndex        =   3
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lblLackThorns 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Tree Lacks Thorns"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label LblThorn 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Short Trees with Thorns"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label LblPetless 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Petioles shorter, never 1 1/2 in length"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3360
      Width           =   5295
   End
End
Attribute VB_Name = "frmSimpleToothed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmBroadleaves(frmNonTwoRanked.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a form allows for the final identification of a variety of tree genera
Option Explicit

Private Sub cmdEndST_Click()
    End
End Sub

Private Sub cmdReturnST_Click()
    frmSimpleToothed.Visible = False
    frmMinnesotaTrees.Visible = True
End Sub

Private Sub imgBass_Click()
    imgBass2.Visible = True
    imgHawthorne.Visible = False
    imgPlums.Visible = False
    imgPoplar.Visible = False
    imgBirch.Visible = False
    imgMoreToothedLeaves.Visible = False
    imgBass.Visible = False
    lblBirch.Visible = False
    lblPlums.Visible = False
    lblNotPaper.Visible = False
    lblPoplar.Visible = False
    lblBass.Visible = False
    LblPetless.Visible = False
    lblPetmore.Visible = False
    lblLackThorns.Visible = False
    lblHaw.Visible = False
    LblThorn.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Tilia, and is commonly known as a Basswood or Linder", , "Genus: Tilia"
    imgBass2.Visible = False
    imgHawthorne.Visible = True
    imgPlums.Visible = True
    imgPoplar.Visible = True
    imgBirch.Visible = True
    imgMoreToothedLeaves.Visible = True
    imgBass.Visible = True
    lblBirch.Visible = True
    lblPlums.Visible = True
    lblNotPaper.Visible = True
    lblPoplar.Visible = True
    lblBass.Visible = True
    LblPetless.Visible = True
    lblPetmore.Visible = True
    lblLackThorns.Visible = True
    lblHaw.Visible = True
    LblThorn.Visible = True
End Sub

Private Sub imgBirch_Click()
    imgBirch2.Visible = True
    imgHawthorne.Visible = False
    imgPlums.Visible = False
    imgPoplar.Visible = False
    imgBirch.Visible = False
    imgMoreToothedLeaves.Visible = False
    imgBass.Visible = False
    lblBirch.Visible = False
    lblPlums.Visible = False
    lblNotPaper.Visible = False
    lblPoplar.Visible = False
    lblBass.Visible = False
    LblPetless.Visible = False
    lblPetmore.Visible = False
    lblLackThorns.Visible = False
    lblHaw.Visible = False
    LblThorn.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Betula, and is commonly known as a Birch ", , "Genus: Betula"
    imgBirch2.Visible = False
    imgHawthorne.Visible = True
    imgPlums.Visible = True
    imgPoplar.Visible = True
    imgBirch.Visible = True
    imgMoreToothedLeaves.Visible = True
    imgBass.Visible = True
    lblBirch.Visible = True
    lblPlums.Visible = True
    lblNotPaper.Visible = True
    lblPoplar.Visible = True
    lblBass.Visible = True
    LblPetless.Visible = True
    lblPetmore.Visible = True
    lblLackThorns.Visible = True
    lblHaw.Visible = True
    LblThorn.Visible = True
    'This long series of Visible Trues and Falses allows the viewer to click the button
End Sub
Private Sub imgHawthorne_Click()
    imgHaw2.Visible = True
    imgHawthorne.Visible = False
    imgPlums.Visible = False
    imgPoplar.Visible = False
    imgBirch.Visible = False
    imgMoreToothedLeaves.Visible = False
    imgBass.Visible = False
    lblBirch.Visible = False
    lblPlums.Visible = False
    lblNotPaper.Visible = False
    lblPoplar.Visible = False
    lblBass.Visible = False
    LblPetless.Visible = False
    lblPetmore.Visible = False
    lblLackThorns.Visible = False
    lblHaw.Visible = False
    LblThorn.Visible = False
     MsgBox "Your tree is a deciduous tree in the genus Crataegus, and is commonly known as a Hawthorne ", , "Genus: Crataegus"
    imgHaw2.Visible = False
    imgHawthorne.Visible = True
    imgPlums.Visible = True
    imgPoplar.Visible = True
    imgBirch.Visible = True
    imgMoreToothedLeaves.Visible = True
    imgBass.Visible = True
    lblBirch.Visible = True
    lblPlums.Visible = True
    lblNotPaper.Visible = True
    lblPoplar.Visible = True
    lblBass.Visible = True
    LblPetless.Visible = True
    lblPetmore.Visible = True
    lblLackThorns.Visible = True
    lblHaw.Visible = True
    LblThorn.Visible = True
End Sub

Private Sub imgMoreToothedLeaves_Click()
    frmSimpleToothed.Hide
    frmToothedRanked.Show
End Sub

Private Sub imgPlums_Click()
    imgPlums2.Visible = True
    imgHawthorne.Visible = False
    imgPlums.Visible = False
    imgPoplar.Visible = False
    imgBirch.Visible = False
    imgMoreToothedLeaves.Visible = False
    imgBass.Visible = False
    lblBirch.Visible = False
    lblPlums.Visible = False
    lblNotPaper.Visible = False
    lblPoplar.Visible = False
    lblBass.Visible = False
    LblPetless.Visible = False
    lblPetmore.Visible = False
    lblLackThorns.Visible = False
    lblHaw.Visible = False
    LblThorn.Visible = False
     MsgBox "Your tree is a deciduous tree in the genus Prunus, and is commonly known as a Plum Tree", , "Genus: Prunus"
    imgPlums2.Visible = False
    imgHawthorne.Visible = True
    imgPlums.Visible = True
    imgPoplar.Visible = True
    imgBirch.Visible = True
    imgMoreToothedLeaves.Visible = True
    imgBass.Visible = True
    lblBirch.Visible = True
    lblPlums.Visible = True
    lblNotPaper.Visible = True
    lblPoplar.Visible = True
    lblBass.Visible = True
    LblPetless.Visible = True
    lblPetmore.Visible = True
    lblLackThorns.Visible = True
    lblHaw.Visible = True
    LblThorn.Visible = True
End Sub

Private Sub imgPoplar_Click()
    imgPoplars.Visible = True
    imgHawthorne.Visible = False
    imgPlums.Visible = False
    imgPoplar.Visible = False
    imgBirch.Visible = False
    imgMoreToothedLeaves.Visible = False
    imgBass.Visible = False
    lblBirch.Visible = False
    lblPlums.Visible = False
    lblNotPaper.Visible = False
    lblPoplar.Visible = False
    lblBass.Visible = False
    LblPetless.Visible = False
    lblPetmore.Visible = False
    lblLackThorns.Visible = False
    lblHaw.Visible = False
    LblThorn.Visible = False
     MsgBox "Your tree is a deciduous tree in the genus Populus, and is commonly known as Populars, Cottonwoods or Aspens", , "Genus: Populus"
    imgPoplars.Visible = False
    imgHawthorne.Visible = True
    imgPlums.Visible = True
    imgPoplar.Visible = True
    imgBirch.Visible = True
    imgMoreToothedLeaves.Visible = True
    imgBass.Visible = True
    lblBirch.Visible = True
    lblPlums.Visible = True
    lblNotPaper.Visible = True
    lblPoplar.Visible = True
    lblBass.Visible = True
    LblPetless.Visible = True
    lblPetmore.Visible = True
    lblLackThorns.Visible = True
    lblHaw.Visible = True
    LblThorn.Visible = True
End Sub
