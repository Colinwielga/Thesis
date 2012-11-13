VERSION 5.00
Begin VB.Form frmLeaves 
   BackColor       =   &H80000012&
   Caption         =   "Leaf Type"
   ClientHeight    =   6240
   ClientLeft      =   2310
   ClientTop       =   2760
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8580
   Begin VB.CommandButton cmdEndLeaf 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   1
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdBegin 
      BackColor       =   &H8000000E&
      Caption         =   "Go Back to Beginning"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Image imgBroadLeaf 
      Height          =   1470
      Left            =   3600
      Picture         =   "frmLeaves.frx":0000
      Top             =   3240
      Width           =   2025
   End
   Begin VB.Image imgSAleaf 
      Height          =   1035
      Left            =   5640
      Picture         =   "frmLeaves.frx":9C72
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Image imgNeedleLeaf 
      Height          =   1215
      Left            =   1440
      Picture         =   "frmLeaves.frx":F0CC
      Top             =   1320
      Width           =   1740
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "If Scale-like or Awl-like click picture"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "If Needle-like click picture"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "If broader than either of these than click on the picture below"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblConifers 
      BackColor       =   &H80000012&
      Caption         =   "What Type of leaves does your tree have?"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   1440
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3360
      Top             =   4080
      Width           =   2175
   End
End
Attribute VB_Name = "frmLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Identifying and Organizing sets of Trees from Minnesota
'frmLeaves(frmLeaves.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'A transition form that allows for preliminary idetification of the leaves as Broad-leaved,Needle-like, or Scale-like/ Awl-like
Private Sub cmdBegin_Click()
    frmLeaves.Hide
    frmMinnesotaTrees.Show
End Sub

Private Sub cmdEndLeaf_Click()
    End
End Sub
Private Sub imgBroadLeaf_Click()
    frmLeaves.Hide
    'leaves the form frmLeaves
    frmBroadleaves.Show
    'goes to form Broadleaves
End Sub

Private Sub imgNeedleLeaf_Click()
    frmLeaves.Visible = False
    frmNeedleLeaves.Visible = True
End Sub

Private Sub imgSAleaf_Click()
    frmLeaves.Hide
    frmSALeaves.Show
End Sub
