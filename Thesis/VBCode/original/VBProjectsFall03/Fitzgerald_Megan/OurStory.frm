VERSION 5.00
Begin VB.Form frmOurStory 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Our Story (Megan Fitzgerald)"
   ClientHeight    =   6495
   ClientLeft      =   4275
   ClientTop       =   2790
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Homepage"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   """Amen, I say to you, whatever you did for one of the least brothers of mine, you did for me."" - Jesus"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"OurStory.frx":0000
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
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Our Story"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmOurStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectAmigosForChrist (Megan Fitzgerald's Visual Basic Project.vbp)
'Form Name : frmOurStory (OurStory.frm)
'Author: Megan Fitzgerald
'Date Written: October 28, 2003
'Purpose of This Form: The purpose of this form is to allow the user the
                        'opportunity to become more familiar with this non-profit
                        'organization by allowing them to read about how it was started.
                        'By using labels, the author is able to give the user infomation
                        'about what motivates the people of Amigos for Christ to serve God
                        'by serving the people of Nicaragua.

Option Explicit

Private Sub cmdReturn_Click()

'Takes the user back to the Homepage "Amigos for Christ".
frmOurStory.Hide
frmHomepage.Show

End Sub

