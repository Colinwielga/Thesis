VERSION 5.00
Begin VB.Form frmProdMP3 
   BackColor       =   &H00FF0000&
   Caption         =   "MP3 Players"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go to Subtotal"
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   6720
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Height          =   1935
      Left            =   6000
      Picture         =   "frmProdMP3.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      Height          =   1695
      Left            =   3360
      Picture         =   "frmProdMP3.frx":1089
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   480
      Picture         =   "frmProdMP3.frx":1F6D
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   2415
      Left            =   3360
      Picture         =   "frmProdMP3.frx":2D96
      ScaleHeight     =   2355
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   6240
      Picture         =   "frmProdMP3.frx":404E
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   5520
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   480
      Picture         =   "frmProdMP3.frx":4514
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton optSamsung 
      BackColor       =   &H00FF0000&
      Caption         =   "Samsung Napster 20.0GB Digital Audio Player"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.OptionButton optRIO 
      BackColor       =   &H00FF0000&
      Caption         =   "Rio Karma 20.0GB Digital Audio Player"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.OptionButton optRCA 
      BackColor       =   &H00FF0000&
      Caption         =   "RCA Lyra 20.0GB MP3/ MP4 Audio/ Video Player"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.OptionButton optRiver 
      BackColor       =   &H00FF0000&
      Caption         =   "iRiver Digital Audio Player with 20.0GB Hard Drive"
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.OptionButton optArchos 
      BackColor       =   &H00FF0000&
      Caption         =   "ARCHOS 20.0GB MP3/MP4 Audio/Video Player"
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   4680
      Width           =   2535
   End
   Begin VB.OptionButton optPod20 
      BackColor       =   &H00FF0000&
      Caption         =   "Apple® iPod™ Digital Audio Player with 20.0GB Hard Drive"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Choose an MP3 player and go to next:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmProdMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjElectrPlus (Joe Dockendorff's VB Project.vbp)
'Form Name : frmProdMP3 (frmProdMP3.frm)
'Author: Joe Dockendorff
'Date Written: March 13, 2004
'Purpose of Form: This form asks the user to choose an mp3 player
                 'and then click the go to subtotal button. This is
                 'the last product the user is prompted to choose
                 'and therefore goes to the subtotal instead of
                 'going to pick another product.
                 
'Option Explicit is a command to force
'the user to declare all variables
'before they can be used.
Option Explicit

Private Sub cmdGo_Click()

If optPod20 = True Then
    M = 1
ElseIf optRiver = True Then
    M = 2
ElseIf optSamsung = True Then
    M = 3
ElseIf optRCA = True Then
    M = 4
ElseIf optRIO = True Then
    M = 5
ElseIf optArchos = True Then
    M = 6
End If

frmProdMP3.Hide
frmSubtotal.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
ReDim MP3(1 To 6) As String
ReDim MP3Price(1 To 6) As Single
Path = "N:\CS130\handin\Dockendorff, Joe\"

'Open the file associated with the product, in this case, the file
'containing MP3 player information.
Close #1
Open Path & "mp3.txt" For Input As #1

For M = 1 To 6
    Input #1, MP3(M), MP3Price(M)
Next M

cmdGo.Enabled = False
End Sub

Private Sub optArchos_Click()
cmdGo.Enabled = True
End Sub

Private Sub optPod20_Click()
cmdGo.Enabled = True
End Sub

Private Sub optRCA_Click()
cmdGo.Enabled = True
End Sub

Private Sub optRIO_Click()
cmdGo.Enabled = True
End Sub

Private Sub optRiver_Click()
cmdGo.Enabled = True
End Sub

Private Sub optSamsung_Click()
cmdGo.Enabled = True
End Sub
