VERSION 5.00
Begin VB.Form frmMW 
   BackColor       =   &H80000012&
   Caption         =   "Marty Walsh"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H80000012&
      Caption         =   "Return to Profile Page"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   8
      Top             =   6840
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   120
      Picture         =   "frmMW.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Step 3: Profit"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   5160
      TabIndex        =   7
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "Step 2:"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Step 1: Fix Global Warming"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "* Major plan for life:"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "* Favorite Cartoon: Captain Planet"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   5160
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "* Major influences in his life / beard: Bob Ross, Chuck Norris, Santa"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "*Not Actual Photo"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   4815
   End
End
Attribute VB_Name = "frmMW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project
'frmDG (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: To create a  profile page for Marty Walsh that can be linked to the team bio page

Private Sub cmdReturn_Click()
    frmMW.Visible = False
    frmBio.Visible = True
End Sub
