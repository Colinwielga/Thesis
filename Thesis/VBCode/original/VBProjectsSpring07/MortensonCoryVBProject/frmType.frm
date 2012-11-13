VERSION 5.00
Begin VB.Form frmType 
   BackColor       =   &H00004000&
   Caption         =   "Filing Status"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSeperate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Married Filing Seperately"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdHead 
      BackColor       =   &H0000FFFF&
      Caption         =   "Head of Household"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdMarried 
      BackColor       =   &H0000FFFF&
      Caption         =   "Married"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdClaimed 
      BackColor       =   &H0000FFFF&
      Caption         =   "Claimed Dependent"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdSolo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Single"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lbl9 
      BackColor       =   &H00004000&
      Caption         =   "Married Filing Seperately: Each spouse files his or her own return, reporting only his or her income, deductions, and exemptions."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2280
      TabIndex        =   11
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00004000&
      Caption         =   "Maried Filing Jointly: Maried couples who combine their income and allowable deductions and file one tax return."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2280
      TabIndex        =   9
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00004000&
      Caption         =   "Can be claimed by parent or other legal guardian as a dependent."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00004000&
      Caption         =   $"frmType.frx":0000
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00004000&
      Caption         =   $"frmType.frx":0090
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00004000&
      Caption         =   "Choose One Type:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00004000&
      Caption         =   "Filing Status"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClaimed_Click()
    frmType.Hide
    frmClaimed.Show
End Sub

Private Sub cmdHead_Click()
    frmTax1.Hide
    frmHead.Show
End Sub

Private Sub cmdMarried_Click()
    frmType.Hide
    frmMarried.Show
End Sub

Private Sub cmdSeperate_Click()
    frmType.Hide
    frmSeperate.Show
End Sub

Private Sub cmdSolo_Click()
    
    'This moves from first to corresponding slide
    frmType.Hide
    frmSingle.Show
    
End Sub

