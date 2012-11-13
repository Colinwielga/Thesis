VERSION 5.00
Begin VB.Form frmFishform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   Picture         =   "Fishform.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H80000013&
      Caption         =   "Go to Main page"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   4095
      Left            =   1680
      Picture         =   "Fishform.frx":1F1502
      ScaleHeight     =   4035
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats on the purchase of your new fish!"
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
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   9615
   End
End
Attribute VB_Name = "frmFishform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'shows the user the fish they have purchased

Private Sub cmdMain_Click()
frmFishform.Hide
Welcomeform2.Show
End Sub

Private Sub Picture1_Click()

End Sub
