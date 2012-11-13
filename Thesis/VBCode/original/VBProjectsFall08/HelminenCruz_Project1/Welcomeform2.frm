VERSION 5.00
Begin VB.Form Welcomeform2 
   Caption         =   "Form3"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form3"
   Picture         =   "Welcomeform2.frx":0000
   ScaleHeight     =   7365
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   1320
      Picture         =   "Welcomeform2.frx":EBED2
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   3480
      Picture         =   "Welcomeform2.frx":F3C24
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   1320
      Picture         =   "Welcomeform2.frx":FBF36
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   3480
      Picture         =   "Welcomeform2.frx":104240
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Get ID"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdGoStore 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Go Shopping"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdFindPet 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Choose a pet"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton cmdGoTrivia 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCitations 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Citations"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdNames 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Name Options"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblwelcome 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Pet Shop"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "Welcomeform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'also the main page to access buttons after you have made your
'id.

Option Explicit


Private Sub cmdCitations_Click()
frmwelcomeform.Hide
frmcitations.Show
End Sub

Private Sub cmdFindPet_Click()

frmwelcomeform.Hide
frmpetform.Show

End Sub

Private Sub cmdGoStore_Click()

frmwelcomeform.Hide
frmstoreform.Show

End Sub

Private Sub cmdGoTrivia_Click()
Welcomeform2.Hide
Triviaform.Show
End Sub

Private Sub cmdNames_Click()
frmwelcomeform.Hide
Form2.Show
End Sub

Private Sub cmdquit_Click()
End

End Sub

Private Sub Command1_Click()

frmwelcomeform.Hide
frmidform.Show



End Sub

Private Sub Form_Load()
Command1.Visible = False
cmdNames.Visible = True
cmdGoStore.Visible = True
cmdFindPet.Visible = True
cmdCitations.Visible = True
cmdGoTrivia.Visible = True


End Sub

