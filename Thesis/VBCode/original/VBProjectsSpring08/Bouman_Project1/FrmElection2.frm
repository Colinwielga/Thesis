VERSION 5.00
Begin VB.Form FrmElection2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Candidates"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSlogan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   1335
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   480
      Width           =   5175
   End
   Begin VB.PictureBox PicCandidate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   3720
      ScaleHeight     =   4575
      ScaleWidth      =   5415
      TabIndex        =   5
      Top             =   2040
      Width           =   5415
   End
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton CmdClinton 
      BackColor       =   &H00FF0000&
      Caption         =   "Hillary Clinton"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton CmdObama 
      BackColor       =   &H00FF0000&
      Caption         =   "Barack Obama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton CmdMcCain 
      BackColor       =   &H000000FF&
      Caption         =   "John McCain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Candidates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "FrmElection2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'Election Project
    'FrmElection2
    'Ian Bouman
    'Written on 3/13
    'The objective of this form is to display which candidate is
    'which, and to show one of the slogans each candidate is running under.
    'Each button triggers a picture to be loaded in the picture
    'box as well as shows the slogan in another picture box.
Private Sub CmdBack_Click()
FrmElection2.Hide
FrmElection1.Show
End Sub

Private Sub CmdClinton_Click()
PicSlogan.Cls
PicCandidate.Cls
PicSlogan.Print Tab(6); "Make History"
PicCandidate.Picture = LoadPicture(App.Path & "/PicHillaryClinton1.jpg")
End Sub

Private Sub CmdMcCain_Click()
PicSlogan.Cls
PicCandidate.Cls
PicSlogan.Print Tab(0); "Ready From Day One."
PicCandidate.Picture = LoadPicture(App.Path & "/PicJohnMcCain2.jpg")
End Sub

Private Sub CmdObama_Click()
PicSlogan.Cls
PicCandidate.Cls
PicSlogan.Print Tab(7); "Yes we can!"
PicCandidate.Picture = LoadPicture(App.Path & "/PicBarackObama2.jpg")
End Sub

Private Sub CmdQuit_Click()
End
End Sub
