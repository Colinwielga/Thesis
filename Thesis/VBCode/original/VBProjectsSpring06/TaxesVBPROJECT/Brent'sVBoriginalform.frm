VERSION 5.00
Begin VB.Form frmFrontpage 
   BackColor       =   &H80000013&
   Caption         =   "Title Page"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTaxFormGo 
      Caption         =   "Go To Tax Form"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdMission 
      Caption         =   "IRS Mission Statement"
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "Commissioner's Message"
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdhelp 
      Caption         =   "Help"
      Height          =   615
      Left            =   6240
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdPersonal 
      Caption         =   "Personal Information"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdQualify 
      BackColor       =   &H80000013&
      Caption         =   "Do you qualify for 1040EZ?"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblMyName 
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Brent Mergen and Associates"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   $"BRENT'~1.frx":0000
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   7815
   End
   Begin VB.Label Title 
      BackColor       =   &H80000013&
      Caption         =   "Welcome to E - Z Taxes"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmFrontpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Title Page (frmFrontpage)
'Brent Timothy Mergen
'24 March 2006
'this is the Title Page, it is where you start the program and return to when finished with your tax form.

Private Sub cmdhelp_Click()
    frmFrontpage.Hide 'hides old form
    frmHelp.Show 'brings you to a new form
End Sub

Private Sub cmdMessage_Click()
    frmFrontpage.Hide 'hides old form
    frmMessage.Show 'brings you to a new form
End Sub

Private Sub cmdMission_Click()
    MsgBox "Provide America's taxpayers top quality service by helping them understand and meet their tax responsibilities and by applying the tax law with integrity and fairness to all.", , "IRS Mission Statement"
    frmFrontpage.Hide 'hides old form
    frmMessage.Show 'brings you to a new form
End Sub

Private Sub cmdPersonal_Click()
    frmFrontpage.Hide 'hides old form
    frmPersonalInfo.Show 'brings you to a new form
End Sub

Private Sub cmdQualify_Click()
    frmFrontpage.Hide 'hides old form
    frmQualify.Show 'brings you to a new form
End Sub

Private Sub cmdTaxFormGo_Click()
    frmFrontpage.Hide 'hides old form
    frmTaxInput.Show 'brings you to a new form
End Sub
