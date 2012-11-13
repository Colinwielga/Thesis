VERSION 5.00
Begin VB.Form frmAirborne 
   BackColor       =   &H00FFFF80&
   Caption         =   "US Army Airborne School"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19050
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   19050
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   9960
      ScaleHeight     =   1455
      ScaleWidth      =   7455
      TabIndex        =   6
      Top             =   1440
      Width           =   7455
   End
   Begin VB.CommandButton cmdWhy 
      Caption         =   "Why Did I Go Here?"
      Height          =   855
      Left            =   10560
      TabIndex        =   5
      Top             =   9120
      Width           =   1815
   End
   Begin VB.PictureBox picAirborne2 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   360
      ScaleHeight     =   9135
      ScaleWidth      =   8895
      TabIndex        =   4
      Top             =   1320
      Width           =   8895
   End
   Begin VB.PictureBox picAirborne1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   9960
      ScaleHeight     =   5535
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   3360
      Width           =   7455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Form"
      Height          =   855
      Left            =   13560
      TabIndex        =   1
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit Program"
      Height          =   855
      Left            =   15600
      TabIndex        =   0
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Tim Johnson   3/20    To State Why I went to Airborne and to show a bit about it"
      Height          =   495
      Left            =   15240
      TabIndex        =   7
      Top             =   10560
      Width           =   2895
   End
   Begin VB.Label lblAirborneTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "US ARMY AIRBORNE SCHOOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "frmAirborne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click() 'Goes back to Main Form
frmMain.Show                'Goes back to Main Form
frmAirborne.Hide
End Sub

Private Sub cmdQuit_Click() 'Ends program where you are
    End                     'Ends program where you are
End Sub

Private Sub cmdWhy_Click()  'Answers a simple question

picInfo.Print "I went to the Jump School to get some extra training and to, hopefully, accelerate my Army career."; Tab(3); " Not every gets a chance to go, and not everyone that goes can make it, so for those of us that pass,"
picInfo.Print "   it is one more thing to lift us up."

End Sub

Private Sub Form_Load()     'Puts ups pictures to improve form appearance

picAirborne1.Picture = LoadPicture(App.Path & "\" & airbornepix(2))
picAirborne2.Picture = LoadPicture(App.Path & "\" & airbornepix(1))

End Sub


