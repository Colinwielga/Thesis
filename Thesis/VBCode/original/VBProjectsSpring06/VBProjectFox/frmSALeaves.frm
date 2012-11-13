VERSION 5.00
Begin VB.Form frmSALeaves 
   BackColor       =   &H00004000&
   Caption         =   "Awl and Scale Like Leaves"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEndSA 
      Caption         =   "End Program"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      TabIndex        =   6
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdRedCedar 
      Caption         =   "Click to learn more about the Red Cedar/Juniper"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdWhiteCedar 
      Caption         =   "Click to learn more about the White Cedar/Arbor Vitae"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton cmdReturnSA 
      BackColor       =   &H8000000D&
      Caption         =   "Return to the Beginning of the Program"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image imgJuni 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSALeaves.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image imgArbor 
      Height          =   4500
      Left            =   0
      Picture         =   "frmSALeaves.frx":29E92
      Top             =   360
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblRC 
      BackColor       =   &H00004000&
      Caption         =   "If your tree has scale or awl-like leaves that are not flattened then it is a Red Cedar or other Juniper."
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Image imgRedCedar 
      Height          =   6000
      Left            =   5400
      Picture         =   "frmSALeaves.frx":53D24
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5520
   End
   Begin VB.Label lblCedars 
      BackColor       =   &H00004000&
      Caption         =   "Trees with Scale- or Awl -like Leaves "
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label lblWC 
      BackColor       =   &H00004000&
      Caption         =   "If your tree has scale-like leaves and flattened twigs then it is a White Cedar or Arbor Vitae . "
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Image imgWhiteCedar 
      Enabled         =   0   'False
      Height          =   6015
      Left            =   0
      Picture         =   "frmSALeaves.frx":F0486
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5460
   End
End
Attribute VB_Name = "frmSALeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Identifying and Organizing sets of Trees from Minnesota
'frmSALeaves(frmSALeaves.frm)
'Author: Kelly Fox
'Date Written:3/19/2006
'This form allows for a final identification of the White and Red Cedars
Option Explicit

Private Sub cmdEndSA_Click()
    End
    'Ends program
End Sub

Private Sub cmdRedCedar_Click()
    imgJuni.Visible = True
    imgWhiteCedar.Visible = False
    imgRedCedar.Visible = False
    cmdWhiteCedar.Visible = False
    cmdRedCedar.Visible = False
    lblRC.Visible = False
    lblWC.Visible = False
     MsgBox "Your tree is a coniferous, evergreen tree in the genus Juniperus, and is commonly known as a Red Cedar or other juniper", , "Genus: Juniperus"
    'Describes the Red Cedar Tree
    imgJuni.Visible = False
    imgWhiteCedar.Visible = True
    imgRedCedar.Visible = True
    cmdWhiteCedar.Visible = True
    cmdRedCedar.Visible = True
    lblRC.Visible = True
    lblWC.Visible = True
End Sub

Private Sub cmdReturnSA_Click()
    frmSALeaves.Hide
    'Leaves the form frmSALeaves
    frmMinnesotaTrees.Show
    'Goes to form frmMinnesotaTrees
End Sub

Private Sub cmdWhiteCedar_Click()
    imgArbor.Visible = True
    imgWhiteCedar.Visible = False
    imgRedCedar.Visible = False
    cmdWhiteCedar.Visible = False
    cmdRedCedar.Visible = False
    lblRC.Visible = False
    lblWC.Visible = False
    MsgBox "Your tree is a coniferous, evergreen in the genus Thuja, and is commmonly known as a White Cedar or Arbor Vitae", , "Genus: Thuja"
    'Describes the White Cedar Tree
    imgArbor.Visible = False
    imgWhiteCedar.Visible = True
    imgRedCedar.Visible = True
    cmdWhiteCedar.Visible = True
    cmdRedCedar.Visible = True
    lblRC.Visible = True
    lblWC.Visible = True
End Sub

