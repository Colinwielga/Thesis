VERSION 5.00
Begin VB.Form frmBroadleaves 
   Caption         =   "Simple or Compound Leaves"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGotoCompound 
      Caption         =   "Click Here "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdSimTooth 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdGotoSimple 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdEndBroadLeaves 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   1
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturnBroad 
      Caption         =   "None of These: Click Here to Return to Beginning"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblGoToSimpleToothed 
      BackStyle       =   0  'Transparent
      Caption         =   "Simple and toothed leaves:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Compound(Have more than one leaflet per leaf) leaves: "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   5520
      TabIndex        =   4
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label lblGotoLobed 
      BackStyle       =   0  'Transparent
      Caption         =   " Simple (have one leaflet per petiole) and lobed leaves: "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblTitleBroad 
      BackStyle       =   0  'Transparent
      Caption         =   "If your trees have: "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image imgSimpTooth 
      Height          =   3690
      Left            =   0
      Picture         =   "frmBroadleaves.frx":0000
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   4635
   End
   Begin VB.Image imgSimpLobe 
      Height          =   3495
      Left            =   0
      Picture         =   "frmBroadleaves.frx":BDDC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4755
   End
   Begin VB.Image imgComp 
      Height          =   8160
      Left            =   4680
      Picture         =   "frmBroadleaves.frx":197D84
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   6360
   End
End
Attribute VB_Name = "frmBroadleaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmBroadleaves(frmBroadLeaves.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This is a transition form to separate the simple toothed and lobed leaves from the compound leaves
Option Explicit


Private Sub cmdEndBroadLeaves_Click()
    End
End Sub

Private Sub cmdGotoCompound_Click()
    'Brings user to form to identify compound leaves
    frmBroadleaves.Hide
    frmCompound.Show
End Sub

Private Sub cmdGotoSimple_Click()
    'Brings user to form to identify broad leaves
    frmBroadleaves.Hide
    frmSimple.Show
End Sub

Private Sub cmdReturnBroad_Click()
    'Brings user back to frmMinnesotaTrees form
    frmBroadleaves.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub cmdSimTooth_Click()
    'Brings user to form wher to identify trees with leaves with teeth
    frmBroadleaves.Hide
    frmSimpleToothed.Show
End Sub

