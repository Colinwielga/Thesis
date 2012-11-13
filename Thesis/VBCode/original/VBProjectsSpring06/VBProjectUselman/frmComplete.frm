VERSION 5.00
Begin VB.Form frmComplete 
   Caption         =   "Complete"
   ClientHeight    =   6750
   ClientLeft      =   1815
   ClientTop       =   2145
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   Picture         =   "frmComplete.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   12135
   Begin VB.PictureBox picWolf 
      Height          =   11310
      Left            =   0
      Picture         =   "frmComplete.frx":160F7
      ScaleHeight     =   11250
      ScaleWidth      =   12120
      TabIndex        =   0
      Top             =   0
      Width           =   12180
      Begin VB.CommandButton cmdQ4 
         Caption         =   "Go Back To Main Menu"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         MaskColor       =   &H8000000F&
         Picture         =   "frmComplete.frx":2C1EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label lblMe 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lance Uselman"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label lblComplete1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have finished the Wildlife Challenge! "
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   5175
      End
      Begin VB.Label lblComplete 
         BackStyle       =   0  'Transparent
         Caption         =   "Congratulations!"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wildlife Challenge (Project1.vbp)
'frmComplete (frmComplete.frm)
'Lance Uselman
'March 24, 2006
'Purpose of the form: The purpose of this form is to congratulate the user on
'                     completion of the quiz.

Option Explicit
Private Sub cmdQ4_Click()
    frmMain.Show
    frmComplete.Hide    'This button allows the user to go back to the main form.
End Sub
