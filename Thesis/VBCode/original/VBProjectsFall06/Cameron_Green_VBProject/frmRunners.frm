VERSION 5.00
Begin VB.Form frmRunners 
   BackColor       =   &H00008000&
   Caption         =   "Legendary Runners Page"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Homepage"
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdKipketer 
      BackColor       =   &H000000FF&
      Caption         =   "Wilson Kipketer"
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdLasse 
      BackColor       =   &H000000FF&
      Caption         =   "Lasse Viren"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdKomen 
      BackColor       =   &H000000FF&
      Caption         =   "Daniel Komen"
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdKennedy 
      BackColor       =   &H000000FF&
      Caption         =   "Bob Kennedy"
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdHicham 
      BackColor       =   &H000000FF&
      Caption         =   "Hicham El Guerrouj"
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdCoe 
      BackColor       =   &H000000FF&
      Caption         =   "Sebastian Coe"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdSnell 
      BackColor       =   &H000000FF&
      Caption         =   "Peter Snell"
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdRyun 
      BackColor       =   &H000000FF&
      Caption         =   "Jim Ryun"
      Height          =   975
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrefontaine 
      BackColor       =   &H000000FF&
      Caption         =   "Steve Prefontaine"
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdBekele 
      BackColor       =   &H000000FF&
      Caption         =   "Kenenisa Bekele"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdZatopek 
      BackColor       =   &H000000FF&
      Caption         =   "Emil Zatopek"
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdTergat 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Tergat"
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdHaile 
      BackColor       =   &H000000FF&
      Caption         =   "Haile Gebrsellasie"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblBiography 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Biography Information on Many Legendary Distance Runners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1575
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmRunners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this button goes back to the homepage from the runner information page'
Private Sub cmdBack_Click()
    frmRunners.Hide
    frmIntroCC.Show
End Sub

'this button goes from the runner information page to the Kenenisa Bekele page'
Private Sub cmdBekele_Click()
    frmRunners.Hide
    frmInfoBekele.Show
End Sub

'this button goes from the runner information page to the Sebastian Coe page'
Private Sub cmdCoe_Click()
    frmRunners.Hide
    frmInfoCoe.Show
End Sub

'this button goes from the runner information page to the Haile Gebrsellasie page'
Private Sub cmdHaile_Click()
    frmRunners.Hide
    frmInfoHaile.Show
End Sub

'this button goes from the runner information page to the Hicham El Guerrouj page'
Private Sub cmdHicham_Click()
    frmRunners.Hide
    frmInfoHicham.Show
End Sub

'this button goes from the runner information page to the Bob Kennedy page'
Private Sub cmdKennedy_Click()
    frmRunners.Hide
    frmInfoKennedy.Show
End Sub

'this button goes from the runner information page to the Wilson Kipketer page'
Private Sub cmdKipketer_Click()
    frmRunners.Hide
    frmInfoKipketer.Show
End Sub

'this button goes from the runner information page to the Daniel Komen page'
Private Sub cmdKomen_Click()
    frmRunners.Hide
    frmInfoKomen.Show
End Sub

'this button goes from the runner information page to the Lasse Viren page'
Private Sub cmdLasse_Click()
    frmRunners.Hide
    frmInfoViren.Show
End Sub

'this button goes from the runner information page to the Steve Prefontaine page'
Private Sub cmdPrefontaine_Click()
    frmRunners.Hide
    frmInfoPre.Show
End Sub

'this button goes from the runner information page to the Jim Ryun page'
Private Sub cmdRyun_Click()
    frmRunners.Hide
    frmInfoRyun.Show
End Sub

'this button goes from the runner information page to the Peter Snell page'
Private Sub cmdSnell_Click()
    frmRunners.Hide
    frmInfoSnell.Show
End Sub

'this button goes from the runner information page to the Paul Tergat page'
Private Sub cmdTergat_Click()
    frmRunners.Hide
    frmInfoTergat.Show
End Sub

'this button goes from the runner information page to the Emil Zatopek page'
Private Sub cmdZatopek_Click()
    frmRunners.Hide
    frmInfoZatopek.Show
End Sub
