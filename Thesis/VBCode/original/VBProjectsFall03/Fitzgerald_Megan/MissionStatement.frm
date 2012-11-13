VERSION 5.00
Begin VB.Form frmMissionStatement 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Mission Statement (Megan Fitzgerald)"
   ClientHeight    =   6270
   ClientLeft      =   4065
   ClientTop       =   3015
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "Century Schoolbook"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   8580
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.PictureBox imgBaby 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      Picture         =   "MissionStatement.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"MissionStatement.frx":1BAC2
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"MissionStatement.frx":1BC8D
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Our Mission "
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmMissionStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmMissionStatement (MissionStatement.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: 'The purpose of this form is to offer the user the opportunity to
                        'read the mission statement of Amigos for Christ.
                        'This form uses labels to display this information.

Option Explicit
Private Sub cmdReturn_Click()

'Take the user back to the Homepage "Amigos for Christ"
frmMissionStatement.Hide
frmHomepage.Show


End Sub
