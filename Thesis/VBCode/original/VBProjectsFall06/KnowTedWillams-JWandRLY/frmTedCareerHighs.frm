VERSION 5.00
Begin VB.Form frmTedCareerHighs 
   Caption         =   "Ted Williams' Career Highs"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmTedCareerHighs.frx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Click For Career Highlights"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   10695
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   14955
      TabIndex        =   1
      Top             =   1680
      Width           =   15015
   End
   Begin VB.CommandButton cmdTedCareerHighsBack 
      BackColor       =   &H000000FF&
      Caption         =   "Click to go back to Contents"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTedCareerHighs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTedCareerHighsBack_Click()
    frmTedmenu.Show
    frmTedCareerHighs.Hide
End Sub

Private Sub Command1_Click()
picresults.Print "2292 Games  7706 At-Bats  1798 Runs  2654 Hits  525 2B  71 3B  521 HR  1839 RBIs  2021 Walks  709 Strikeouts  .344 Career Avg. "

 


End Sub
