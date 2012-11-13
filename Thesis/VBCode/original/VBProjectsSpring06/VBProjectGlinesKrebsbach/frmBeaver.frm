VERSION 5.00
Begin VB.Form frmBeaver 
   Caption         =   "Beaver Creek"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   Picture         =   "frmBeaver.frx":0000
   ScaleHeight     =   10185
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Resorts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton cmdtickets 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lift Tickets"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdLodge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lodging"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Airfare"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quick Facts"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Beaver Creek"
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   68.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmBeaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmBeaver(frmBeaver.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form:  This form allows the suser to navigate all the features of the
'beaver creek resort. it allows the user to access forms that search for ticket prices,
'resorts, and airfair. this form also allows the user to read up on quick facts about the
'ski resort
Private Sub cmdAir_Click()
frmBeaver.Hide
frmAirline.Show

End Sub

Private Sub cmdback_Click()
frmBeaver.Visible = False
frmContents.Visible = True

End Sub

Private Sub cmdFacts_Click()
MsgBox "Beaver Creek is a feast for the eyes; a delight for the senses. The influence of renowned resorts such as Switzerland's St. Moritz, Italy's Cortina, and Spain's Val d'Aran has resulted in a unique combination of mountain excitement and village luxury. Beaver Creek Mountain was originally designed to accommodate skiers of all ability levels.  That design proved timeless and today the mountain is enjoyed by a variety of both winter and summer sports enthusiasts.", , "Beaver Creek Facts"

End Sub



Private Sub cmdLodge_Click()
frmBeaver.Hide
frmBCLodge.Show

End Sub

Private Sub cmdtickets_Click()
frmBeaver.Hide
frmBCtickets.Show
End Sub
