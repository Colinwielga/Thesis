VERSION 5.00
Begin VB.Form frmDidKnow 
   Caption         =   "Did you know???"
   ClientHeight    =   12264
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   16932
   LinkTopic       =   "Form1"
   Picture         =   "frmDidKnow.frx":0000
   ScaleHeight     =   12264
   ScaleWidth      =   16932
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContents 
      Caption         =   "Back to the Table of Contents"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   13560
      TabIndex        =   0
      Top             =   10560
      Width           =   2892
   End
   Begin VB.Label lblKnow 
      BackStyle       =   0  'Transparent
      Caption         =   "DID YOU KNOW..."
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   15732
   End
   Begin VB.Label lblHome 
      BackStyle       =   0  'Transparent
      Caption         =   "Arthur Blank, Co-founder of Home Depot and Owner of the Atlanta Falcons is a CPA?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1692
      Left            =   1440
      TabIndex        =   4
      Top             =   8400
      Width           =   14532
   End
   Begin VB.Label lblJP 
      BackStyle       =   0  'Transparent
      Caption         =   " J.P. Morgan’s First Job was a Junior Accountant?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   3
      Top             =   6360
      Width           =   13092
   End
   Begin VB.Label lblFBI 
      BackStyle       =   0  'Transparent
      Caption         =   " Thomas Pickard, FBI’s # 2 in Charge is a CPA?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3000
      TabIndex        =   2
      Top             =   4320
      Width           =   13092
   End
   Begin VB.Label lblKnight 
      BackStyle       =   0  'Transparent
      Caption         =   " Phil Knight, Founder and Chair of Nike is a CPA?"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   13092
   End
End
Attribute VB_Name = "frmDidKnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Did You Know Form
'Tony McLean
'3.31.2008
'The purpose of this form is to allow the user
'to have some fun in learning some interesting facts
'about leaders in our society who come from accounting
'backgrounds.
Private Sub cmdContents_Click()
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Show
    frmIntroduction.Hide
End Sub

Private Sub Label3_Click()

End Sub
