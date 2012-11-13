VERSION 5.00
Begin VB.Form frmwelcomeform 
   BackColor       =   &H00FF00FF&
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   Picture         =   "HomePage.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
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
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
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
      Left            =   6120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   2175
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Get ID"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   2880
      Picture         =   "HomePage.frx":EBED2
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1215
      Left            =   720
      Picture         =   "HomePage.frx":F42C4
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   2880
      Picture         =   "HomePage.frx":FC5CE
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   720
      Picture         =   "HomePage.frx":1048E0
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmwelcomeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'main form to access all buttons to different options and also
'get an id to use the program.



Private Sub cmdFindPet_Click()

frmwelcomeform.Hide
frmpetform.Show

End Sub

Private Sub cmdGoStore_Click()

frmwelcomeform.Hide
frmstoreform.Show

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
Command1.Visible = True
cmdNames.Visible = False
cmdGoStore.Visible = False
cmdFindPet.Visible = False
cmdCitations.Visible = False
cmdGoTrivia.Visible = False

End Sub
