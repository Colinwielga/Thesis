VERSION 5.00
Begin VB.Form frmProdTVs 
   BackColor       =   &H00008000&
   Caption         =   "Televisions"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go to next"
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   6480
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   5520
      Picture         =   "frmProdTVs.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   1080
      Picture         =   "frmProdTVs.frx":0E1F
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   5640
      Picture         =   "frmProdTVs.frx":1FA9
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   1320
      Picture         =   "frmProdTVs.frx":3509
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.OptionButton optPioneer 
      BackColor       =   &H00008000&
      Caption         =   "Pioneer 50"" 16:9 Widescreen HD-Ready PureVision Plasma TV with PureDrive - Black "
      Height          =   975
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   2775
   End
   Begin VB.OptionButton optPhilips 
      BackColor       =   &H00008000&
      Caption         =   "Philips 30"" Widescreen HD-Ready Flat-Tube TV with Active Control"
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   3960
      Width           =   2535
   End
   Begin VB.OptionButton optSamsung 
      BackColor       =   &H00008000&
      Caption         =   "Samsung 50"" Widescreen HD-Ready DLP-Projection TV w/ DVI Input & 2-Tuner PIP"
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.OptionButton optSansui 
      BackColor       =   &H00008000&
      Caption         =   "Sansui 20"" Stereo TV/Hi-Fi VCR/DVD Player Combo"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Please choose a TV and go to next:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmProdTVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjElectrPlus (Joe Dockendorff's VB Project.vbp)
'Form Name : frmProdTVs (frmProdTVs.frm)
'Author: Joe Dockendorff
'Date Written: March 13, 2004
'Purpose of Form: To get user to pick a product they would like
                 'to shop for and then give them a choice of models.
                 'The user can then pick the model of choice and
                 'add the product to their cart and checkout or
                 'shop around some more.  When the user is done, the
                 'total price is displayed.
                 
'Option Explicit is a command to force
'the user to declare all variables
'before they can be used.
Option Explicit

Private Sub cmdGo_Click()
'Determine which option is chosen

If optSansui = True Then
    T = 1
ElseIf optSamsung = True Then
    T = 2
ElseIf optPhilips = True Then
    T = 3
ElseIf optPioneer = True Then
    T = 4
End If

frmProdTVs.Hide
frmStart.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
ReDim TV(1 To 4) As String
ReDim TVPrice(1 To 4) As Single
Path = "N:\CS130\handin\Dockendorff, Joe\"

'Open the file associated with the product, in this case, the file
'containing TV information.
Close #1
Open Path & "TVs.txt" For Input As #1

For T = 1 To 4
    Input #1, TV(T), TVPrice(T)
Next T

cmdGo.Enabled = False
End Sub

Private Sub optPhFlatTube_Click()
cmdGo.Enabled = True
End Sub

Private Sub optPioneer_Click()
cmdGo.Enabled = True
End Sub

Private Sub optSamsung_Click()
cmdGo.Enabled = True
End Sub

Private Sub optSansui_Click()
cmdGo.Enabled = True
End Sub

