VERSION 5.00
Begin VB.Form frmSeparate 
   Caption         =   "Trees with Separated Needles "
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndSepNeedles 
      Caption         =   "End Progam"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   7
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdHemlock 
      Caption         =   "Click Here for more information on Hemlock"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdDouglasFir 
      Caption         =   "Click Here for more information on Douglas Firs"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdFirs 
      Caption         =   "Click Here for more information on Firs"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cmdSpruce 
      Caption         =   "Click Here for more information on Spruces"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdReturnSeparate 
      Caption         =   "Return to Beginning "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Image imgHemlock2 
      Height          =   5880
      Left            =   0
      Picture         =   "frmSeparate.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Image imgDoug2 
      Height          =   9000
      Left            =   0
      Picture         =   "frmSeparate.frx":C6102
      Top             =   -120
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.Image imgFir2 
      Height          =   5580
      Left            =   360
      Picture         =   "frmSeparate.frx":14E9A4
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label LblDoug 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Needles an inch or more in length, blunt (Douglas Fir)"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6120
      TabIndex        =   11
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label lblHem 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Needles less than an inch long, blunt at tip underside silvery (Hemlock)"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   480
      TabIndex        =   10
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label lblFir 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Needles flat and do not roll easily and are sessile without petioles (Firs)"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   6600
      TabIndex        =   9
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblSpruces 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Needles feel angular when rolled between thumb and forefinger (Spruces)"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   3495
   End
   Begin VB.Image imgSpruce2 
      Height          =   4680
      Left            =   120
      Picture         =   "frmSeparate.frx":206B36
      Top             =   1680
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label lblNoPet 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Needles with petioles and do not roll easily:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   1
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Image imgDoug 
      Height          =   4800
      Left            =   5280
      Picture         =   "frmSeparate.frx":235558
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   5520
   End
   Begin VB.Image imgHem 
      Height          =   4680
      Left            =   0
      Picture         =   "frmSeparate.frx":47559A
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   5280
   End
   Begin VB.Label lblTitleSepNeedles 
      BackColor       =   &H00004000&
      Caption         =   "Trees with Separated Needles"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image imgFirs 
      Height          =   4275
      Left            =   5280
      Picture         =   "frmSeparate.frx":6B55DC
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   5610
   End
   Begin VB.Image imgSpruce 
      Height          =   4140
      Left            =   0
      Picture         =   "frmSeparate.frx":6D6AA2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5265
   End
End
Attribute VB_Name = "frmSeparate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmSeparate(frmSeparate.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a form allows for a closer examination of needles that are not in bundles
Option Explicit

Private Sub cmdDouglasFir_Click()
    imgDoug2.Visible = True
    imgSpruce.Visible = False
    imgFirs.Visible = False
    imgHem.Visible = False
    imgDoug.Visible = False
    cmdSpruce.Visible = False
    cmdFirs.Visible = False
    cmdDouglasFir.Visible = False
    cmdHemlock.Visible = False
    lblNoPet.Visible = False
    lblTitleSepNeedles.Visible = False
    lblSpruces.Visible = False
    lblFir.Visible = False
    lblHem.Visible = False
    LblDoug.Visible = False
    MsgBox " Your tree is a coniferous, evergreen in the genus Pseudotsuga, and is commonly known as a Douglas Fir ", , "Genus: Pseudotsuga"
    imgDoug2.Visible = False
    imgSpruce.Visible = True
    imgFirs.Visible = True
    imgHem.Visible = True
    imgDoug.Visible = True
    cmdSpruce.Visible = True
    cmdFirs.Visible = True
    cmdDouglasFir.Visible = True
    cmdHemlock.Visible = True
    lblNoPet.Visible = True
    lblTitleSepNeedles.Visible = True
    lblSpruces.Visible = True
    lblFir.Visible = True
    lblHem.Visible = True
    LblDoug.Visible = True
End Sub

Private Sub cmdEndSepNeedles_Click()
    End
End Sub

Private Sub cmdGotoPicfromSA_Click()
    frmSeparate.Hide
    frmPicTrees.Show
End Sub

Private Sub cmdFirs_Click()
    imgFir2.Visible = True
    imgFirs.Visible = False
    imgHem.Visible = False
    imgDoug.Visible = False
    imgSpruce.Visible = False
    cmdSpruce.Visible = False
    cmdFirs.Visible = False
    cmdDouglasFir.Visible = False
    cmdHemlock.Visible = False
    lblNoPet.Visible = False
    lblTitleSepNeedles.Visible = False
    lblSpruces.Visible = False
    lblFir.Visible = False
    lblHem.Visible = False
    LblDoug.Visible = False
    MsgBox " Your tree is a coniferous, evergreen in the genus Abies, and is commonly known as a Fir ", , "Genus: Abies"
    imgFir2.Visible = False
    imgSpruce.Visible = True
    imgFirs.Visible = True
    imgHem.Visible = True
    imgDoug.Visible = True
    cmdSpruce.Visible = True
    cmdFirs.Visible = True
    cmdDouglasFir.Visible = True
    cmdHemlock.Visible = True
    lblNoPet.Visible = True
    lblTitleSepNeedles.Visible = True
    lblSpruces.Visible = True
    lblFir.Visible = True
    lblHem.Visible = True
    LblDoug.Visible = True
End Sub

Private Sub cmdHemlock_Click()
    imgHemlock2.Visible = True
    imgSpruce.Visible = False
    imgFirs.Visible = False
    imgHem.Visible = False
    imgDoug.Visible = False
    cmdSpruce.Visible = False
    cmdFirs.Visible = False
    cmdDouglasFir.Visible = False
    cmdHemlock.Visible = False
    lblNoPet.Visible = False
    lblTitleSepNeedles.Visible = False
    lblSpruces.Visible = False
    lblFir.Visible = False
    lblHem.Visible = False
    LblDoug.Visible = False
    MsgBox " Your tree is a coniferous, evergreen in the genus Tsuga, and is commonly known as a Hemlock ", , "Genus: Tsuga"
    imgHemlock2.Visible = False
    imgSpruce.Visible = True
    imgFirs.Visible = True
    imgHem.Visible = True
    imgDoug.Visible = True
    cmdSpruce.Visible = True
    cmdFirs.Visible = True
    cmdDouglasFir.Visible = True
    cmdHemlock.Visible = True
    lblNoPet.Visible = True
    lblTitleSepNeedles.Visible = True
    lblSpruces.Visible = True
    lblFir.Visible = True
    lblHem.Visible = True
    LblDoug.Visible = True
End Sub

Private Sub cmdReturnSeparate_Click()
    frmSeparate.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub cmdSpruce_Click()
    imgSpruce2.Visible = True
    imgSpruce.Visible = False
    imgFirs.Visible = False
    imgHem.Visible = False
    imgDoug.Visible = False
    cmdSpruce.Visible = False
    cmdFirs.Visible = False
    cmdDouglasFir.Visible = False
    cmdHemlock.Visible = False
    lblNoPet.Visible = False
    lblTitleSepNeedles.Visible = False
    lblSpruces.Visible = False
    lblFir.Visible = False
    lblHem.Visible = False
    LblDoug.Visible = False
    MsgBox " Your tree is a coniferous, evergreen in the genus Picea, and is commonly known as a Spruce ", , "Genus: Picea"
    imgSpruce2.Visible = False
    imgSpruce.Visible = True
    imgFirs.Visible = True
    imgHem.Visible = True
    imgDoug.Visible = True
    cmdSpruce.Visible = True
    cmdFirs.Visible = True
    cmdDouglasFir.Visible = True
    cmdHemlock.Visible = True
    lblNoPet.Visible = True
    lblTitleSepNeedles.Visible = True
    lblSpruces.Visible = True
    lblFir.Visible = True
    lblHem.Visible = True
    LblDoug.Visible = True
End Sub

