VERSION 5.00
Begin VB.Form frmBasics 
   Caption         =   "SURFING Basics"
   ClientHeight    =   5565
   ClientLeft      =   3975
   ClientTop       =   3105
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   Picture         =   "frmBasics.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   8025
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Text            =   "Click on a button below to view more information."
      Top             =   0
      Width           =   8055
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdsafe 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Safety"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmded 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Surfing Etiquette"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Cmdskills 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Basic Skills"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblintro 
      BackColor       =   &H00400000&
      Caption         =   "This screen will show you the basic information needed to be able to surf."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label lblsafe 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBasics.frx":C35C
      Height          =   3975
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label lbled 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBasics.frx":CB44
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblbasics 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBasics.frx":CF8B
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   7335
   End
End
Attribute VB_Name = "frmBasics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SurfProject (SurfingProject.vbp)
'Form Name: frmBasics (frmDest.frm)
'Author: Benjamin Luther
'Purpose of Form: this form educates the user about the
                    'safety, etiquette, and the basics of surfing
                    'by hiding and displaying different labels
Private Sub cmdclose_Click()
    frmBasics.Hide 'hides the basics form
End Sub

Private Sub cmded_Click()
    lbled.Visible = True 'displays the etiquette label and the information about the proper surfing ways
    lblbasics.Visible = False 'hides the basics label
    lblsafe.Visible = False 'hides the safety label
    lblintro.Visible = False 'hides the intoduction label
End Sub

Private Sub cmdsafe_Click()
    lblsafe.Visible = True 'displays the safety label, showing information about safety issues in surfing
    lbled.Visible = False 'hides the etiquette label
    lblbasics.Visible = False 'hides the basics label
    lblintro.Visible = False 'hides the intoduction label
End Sub

Private Sub Cmdskills_Click()
    lblbasics.Visible = True 'displays the basics label, shows information about the basic skills necessary to surf.
    lbled.Visible = False 'hides the etiquette label
    lblsafe.Visible = False 'hides the safety label
    lblintro.Visible = False 'hides the intoduction label
End Sub


