VERSION 5.00
Begin VB.Form frmTurtleform 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   Picture         =   "Turtleform.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFF80&
      Caption         =   "Go to main page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   4575
      Left            =   720
      Picture         =   "Turtleform.frx":18B902
      ScaleHeight     =   4515
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   3480
      Width           =   6015
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats on the purchase of your new turtle!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   9495
   End
End
Attribute VB_Name = "frmTurtleform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'shows user the turtle they have purchased. Congrats

Private Sub cmdMain_Click()
frmTurtleform.Hide
Welcomeform2.Show

End Sub

Private Sub lblCongrats_Click()

End Sub
