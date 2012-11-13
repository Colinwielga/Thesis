VERSION 5.00
Begin VB.Form frmTrends 
   BackColor       =   &H000000FF&
   Caption         =   "Trends for Winter 2005/2006!"
   ClientHeight    =   9390
   ClientLeft      =   26295
   ClientTop       =   2880
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   Picture         =   "frmTrends.frx":0000
   ScaleHeight     =   9390
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMilitary2 
      Height          =   3015
      Left            =   8520
      Picture         =   "frmTrends.frx":8A54
      ScaleHeight     =   2955
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox picAviator 
      Height          =   3015
      Left            =   6000
      Picture         =   "frmTrends.frx":A8B1
      ScaleHeight     =   2955
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox picMilitary1 
      Height          =   3015
      Left            =   6480
      Picture         =   "frmTrends.frx":BE7E
      ScaleHeight     =   2955
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      MaskColor       =   &H00FFC0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   2655
   End
   Begin VB.PictureBox picVolume 
      Height          =   5775
      Left            =   360
      Picture         =   "frmTrends.frx":DEDF
      ScaleHeight     =   5715
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.PictureBox picVictoriana 
      Height          =   2775
      Left            =   3960
      Picture         =   "frmTrends.frx":11E26
      ScaleHeight     =   2715
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox picRussian 
      Height          =   3375
      Left            =   3960
      Picture         =   "frmTrends.frx":1395A
      ScaleHeight     =   3315
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.PictureBox picSillouette 
      Height          =   3975
      Left            =   8160
      Picture         =   "frmTrends.frx":14E7D
      ScaleHeight     =   3915
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label lblNames 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenna Gebeke ~ Katie Ranallo"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   8400
      Width           =   3015
   End
   Begin VB.Label lblTrends 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Trends for Winter 2005/2006!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "frmTrends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Trends
'Form Objective: This form allows the user to view several trends for Winter 2005/2006.  They are able to click on a picture of a trend and view a message box with information on the trend they selected.

Private Sub cmdReturn_Click()
'This command button allows the user to return to the startup page.
    frmTrends.Hide
    frmStart.Show
End Sub

Private Sub Form_Load()
    frmTrends.Caption = "Welcome " & userName & "  - Trends for Winter 2005/2006!"
End Sub

Private Sub picAviator_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxAviator.Caption = "Not Designed For Heavier-Than-Air Machinery..."
    frmMsgBoxAviator.Visible = True
    frmMsgBoxAviator.BackColor = &H80C0FF
   
End Sub

Private Sub picMilitary1_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxMilitary.Caption = "The Military Was Never So Popular..."
    frmMsgBoxMilitary.Visible = True
    frmMsgBoxMilitary.BackColor = &HC0E0FF
End Sub

Private Sub picMilitary2_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxMilitary.Caption = "The Military Was Never So Popular..."
    frmMsgBoxMilitary.Visible = True
    frmMsgBoxMilitary.BackColor = &HC0E0FF
   
End Sub

Private Sub picRussian_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxRussian.Caption = "Wearing Russian IS Considered Sport!"
    frmMsgBoxRussian.Visible = True
    frmMsgBoxRussian.BackColor = &HC0FFC0
End Sub

Private Sub picSillouette_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxSillhouette.Caption = "Dare To Outline Your Curves..."
    frmMsgBoxSillhouette.Visible = True
    frmMsgBoxSillhouette.BackColor = &HFFC0FF
    
End Sub

Private Sub picVictoriana_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxVictoriana.Caption = "Who Knew We Were Living In A Victorian World?"
    frmMsgBoxVictoriana.Visible = True
    frmMsgBoxVictoriana.BackColor = &HFFFFC0
   
End Sub

Private Sub picVolume_Click()
'By selecting this picture, the user will view a message box describing this style.
    frmMsgBoxVolume.Caption = "Volume = Length x Width x Height"
    frmMsgBoxVolume.Visible = True
    frmMsgBoxVolume.BackColor = &HFFC0C0
   
End Sub
