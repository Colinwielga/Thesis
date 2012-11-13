VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   FillColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Text            =   "2005 Houston Astros"
      Top             =   240
      Width           =   7335
   End
   Begin VB.CommandButton cmdpitching 
      Caption         =   "View Pitching Statistics"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdoffense 
      Caption         =   "View Offensive Statistics"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   3975
      Left            =   2640
      ScaleHeight     =   3915
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label lblstadium 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Minute Maid Park--Home of the Houston Astros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label lbltom 
      BackColor       =   &H000000FF&
      Caption         =   "Tom Wentzell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2005 Houston Astros Statistics(Wentzell_Tom_Project)
'frmmain (frmmain.frm)
'Tom Wentzell
'October 30, 2005
'The purpose of this project is to view individual statistics for the 2005 Houstons Astros.
'The user can view both offensive statistics and pitching statistics, and can sort the
'statistics to view which players lead the team in different statistical categories.
'There is also an opportunity for the user to input his or her own batting statistics to
'calculate a total batting average.
'The purpose of this particular form is to serve as a welcome page to the user and direct
'the user to either an offensive statistics form or a pitching statistics form.

'This loads the picture of Minute Maid Park every time the form is brought up.
Private Sub Form_Load()
picResults.Picture = LoadPicture(App.Path & "\Stadium.jpg")
End Sub

'This command button directs the user to the offensive statistics form.
Private Sub cmdoffense_Click()
    frmoffense.Show
    frmmain.Hide
End Sub

'This command button directs the user to the pitching statistics form.
Private Sub cmdpitching_Click()
    frmpitching.Show
    frmmain.Hide
End Sub

'This command button allows the user to exit the program.
Private Sub cmdquit_Click()
End
End Sub

