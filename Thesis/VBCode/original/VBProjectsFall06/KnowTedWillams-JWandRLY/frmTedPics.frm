VERSION 5.00
Begin VB.Form frmTedPics 
   Caption         =   "Pictures of Ted Williams"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   Picture         =   "frmTedPics.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdweigh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ted Weigh's Bat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdcolorcard 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Color Card"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton cmdtedpilot 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pilot Ted"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdtedandjoe 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ted and Joe"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdhomerun 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Home Run"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdslide 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Slide!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdlastgame 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Last Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton cmdpic1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warm Up"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5355
      ScaleWidth      =   7275
      TabIndex        =   1
      Top             =   3000
      Width           =   7335
   End
   Begin VB.CommandButton cmdTedPicsBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click to go back to Contents"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   8640
      Width           =   7335
   End
End
Attribute VB_Name = "frmTedPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcolorcard_Click()
    picresults.Picture = LoadPicture(App.Path & "\clrcard.jpg")
End Sub

Private Sub cmdhomerun_Click()
    picresults.Picture = LoadPicture(App.Path & "\homerun.jpg")
End Sub

Private Sub cmdlastgame_Click()
    picresults.Picture = LoadPicture(App.Path & "\lastgame.jpg")
End Sub

Private Sub cmdpic1_Click()
    picresults.Picture = LoadPicture(App.Path & "\warmup.jpg")
End Sub

Private Sub cmdslide_Click()
    picresults.Picture = LoadPicture(App.Path & "\slide.jpg")
End Sub

Private Sub cmdtedandjoe_Click()
    picresults.Picture = LoadPicture(App.Path & "\tedandjoe.jpg")
End Sub

Private Sub cmdTedPicsBack_Click()
    frmTedmenu.Show
    frmTedPics.Hide
End Sub

Private Sub cmdtedpilot_Click()
    picresults.Picture = LoadPicture(App.Path & "\pilotted.jpg")
End Sub

Private Sub cmdweigh_Click()
    picresults.Picture = LoadPicture(App.Path & "\weighbat.jpg")
End Sub
