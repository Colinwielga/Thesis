VERSION 5.00
Begin VB.Form frmPinnatelyCompound 
   BackColor       =   &H00404040&
   Caption         =   "Alternate Pinnately Compound Leaved Trees"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHoney 
      Caption         =   "For more info on the Honey Locust: Click Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdMt 
      Caption         =   "For more info on the Mountain Ash: Click Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   6
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdBlLo 
      Caption         =   "For more info on the Black Locust: Click Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturnPA 
      Caption         =   "To Return to the Beginning Press Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdEndAP 
      Caption         =   "To End Program: Click Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdButter 
      Caption         =   "For more info on the Butternut: Click Here"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image imgBLo2 
      Height          =   5130
      Left            =   -120
      Picture         =   "frmPinnatelyCompound.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.Image imgMtAsh2 
      Height          =   3225
      Left            =   0
      Picture         =   "frmPinnatelyCompound.frx":5D32A
      Top             =   -120
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image imgButter2 
      Height          =   5370
      Left            =   0
      Picture         =   "frmPinnatelyCompound.frx":8C748
      Stretch         =   -1  'True
      Top             =   -120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgHoney2 
      Height          =   3705
      Left            =   0
      Picture         =   "frmPinnatelyCompound.frx":139FB2
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Label lblHoney 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPinnatelyCompound.frx":1C3538
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2655
      Left            =   8520
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblMt 
      BackStyle       =   0  'Transparent
      Caption         =   "Lacks thorns, leaves are 6 to 8 inches long; fruits orange-red (Mountain Ash)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3720
      TabIndex        =   9
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label lblBlackLoc 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPinnatelyCompound.frx":1C35CD
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblTitlePA 
      BackStyle       =   0  'Transparent
      Caption         =   "Trees with Alternate, Pinnately Compound Leaves "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label lblButter 
      BackStyle       =   0  'Transparent
      Caption         =   "Lacks thorns and has toothed leaflets with leaves usually 1 foot or more in length (Butternut)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Image imgButternut 
      Height          =   3495
      Left            =   3240
      Picture         =   "frmPinnatelyCompound.frx":1C3683
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3930
   End
   Begin VB.Image imgMountainAsh 
      Height          =   3600
      Left            =   3360
      Picture         =   "frmPinnatelyCompound.frx":208711
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   3840
   End
   Begin VB.Image imgHoneyLocust 
      Height          =   3240
      Left            =   7080
      Picture         =   "frmPinnatelyCompound.frx":240B53
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Image imgBlackLocust 
      Height          =   3585
      Left            =   0
      Picture         =   "frmPinnatelyCompound.frx":29C32D
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "frmPinnatelyCompound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmPinnatelyCompound(frmPinnatelyCompound.frm)
'Author: Kelly Fox
'Date Written:3/23/2006
'This is another identification slide
Option Explicit
Private Sub cmdBlLo_Click()
    imgBLo2.Visible = True
    imgHoneyLocust.Visible = False
    imgButternut.Visible = False
    imgBlackLocust.Visible = False
    imgMountainAsh.Visible = False
    lblHoney.Visible = False
    lblMt.Visible = False
    lblButter.Visible = False
    lblBlackLoc.Visible = False
    cmdHoney.Visible = False
    cmdMt.Visible = False
    cmdBlLo.Visible = False
    cmdButter.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Robinia, and is commonly known as a Black Locust or Common Locust", , "Genus: Robinia"
    imgBLo2.Visible = False
    imgHoneyLocust.Visible = True
    imgButternut.Visible = True
    imgBlackLocust.Visible = True
    imgMountainAsh.Visible = True
    lblHoney.Visible = True
    lblMt.Visible = True
    lblButter.Visible = True
    lblBlackLoc.Visible = True
    cmdHoney.Visible = True
    cmdMt.Visible = True
    cmdBlLo.Visible = True
    cmdButter.Visible = True
End Sub

Private Sub cmdButter_Click()
    imgButter2.Visible = True
    imgHoneyLocust.Visible = False
    imgButternut.Visible = False
    imgBlackLocust.Visible = False
    imgMountainAsh.Visible = False
    lblHoney.Visible = False
    lblMt.Visible = False
    lblButter.Visible = False
    lblBlackLoc.Visible = False
    cmdHoney.Visible = False
    cmdMt.Visible = False
    cmdBlLo.Visible = False
    cmdButter.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Juglans, and is commonly known as a Butternut or Black Walnut ", , "Genus: Juglans"
    imgButter2.Visible = False
    imgHoneyLocust.Visible = True
    imgButternut.Visible = True
    imgBlackLocust.Visible = True
    imgMountainAsh.Visible = True
    lblHoney.Visible = True
    lblMt.Visible = True
    lblButter.Visible = True
    lblBlackLoc.Visible = True
    cmdHoney.Visible = True
    cmdMt.Visible = True
    cmdBlLo.Visible = True
    cmdButter.Visible = True
    
End Sub

Private Sub cmdEndAP_Click()
    End
End Sub

Private Sub cmdHoney_Click()
    imgHoney2.Visible = True
    imgHoneyLocust.Visible = False
    imgButternut.Visible = False
    imgBlackLocust.Visible = False
    imgMountainAsh.Visible = False
    lblHoney.Visible = False
    lblMt.Visible = False
    lblButter.Visible = False
    lblBlackLoc.Visible = False
    cmdHoney.Visible = False
    cmdMt.Visible = False
    cmdBlLo.Visible = False
    cmdButter.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Gleditsia, and is commonly known as a Honey Locust ", , "Genus: Gleditsia"
    imgHoney2.Visible = False
    imgHoneyLocust.Visible = True
    imgButternut.Visible = True
    imgBlackLocust.Visible = True
    imgMountainAsh.Visible = True
    lblBlackLoc.Visible = True
    lblHoney.Visible = True
    lblMt.Visible = True
    lblButter.Visible = True
    cmdHoney.Visible = True
    cmdMt.Visible = True
    cmdBlLo.Visible = True
    cmdButter.Visible = True
End Sub

Private Sub cmdMt_Click()
    imgMtAsh2.Visible = True
    imgHoneyLocust.Visible = False
    imgButternut.Visible = False
    imgBlackLocust.Visible = False
    imgMountainAsh.Visible = False
    lblHoney.Visible = False
    lblMt.Visible = False
    lblButter.Visible = False
    lblBlackLoc.Visible = False
    cmdHoney.Visible = False
    cmdMt.Visible = False
    cmdBlLo.Visible = False
    cmdButter.Visible = False
    MsgBox "Your tree is a deciduous tree in the genus Sorbus, and is commonly known as a Mountain Ash ", , "Genus: Sorbus"
    imgMtAsh2.Visible = False
    imgHoneyLocust.Visible = True
    imgButternut.Visible = True
    imgBlackLocust.Visible = True
    imgMountainAsh.Visible = True
    lblHoney.Visible = True
    lblMt.Visible = True
    lblBlackLoc.Visible = True
    lblButter.Visible = True
    cmdHoney.Visible = True
    cmdMt.Visible = True
    cmdBlLo.Visible = True
    cmdButter.Visible = True
End Sub

Private Sub cmdReturnPA_Click()
    frmPinnatelyCompound.Hide
    frmMinnesotaTrees.Show
End Sub

