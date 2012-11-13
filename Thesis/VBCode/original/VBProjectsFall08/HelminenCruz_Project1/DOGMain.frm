VERSION 5.00
Begin VB.Form frmDogform 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   Picture         =   "DOGMain.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Go to Main Page"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   1680
      Picture         =   "DOGMain.frx":1C963A
      ScaleHeight     =   5475
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats on the purchase of your new dog!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "frmDogform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'main dog page, shows the user the dog they purchased


Private Sub cmdMain_Click()
frmDogform.Hide
Welcomeform2.Show
End Sub

Private Sub Picture1_Click()

End Sub
