VERSION 5.00
Begin VB.Form frmKittenForm 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   Picture         =   "KittenForm.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   6495
      Left            =   -480
      Picture         =   "KittenForm.frx":283F8A
      ScaleHeight     =   6435
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   2040
      Width           =   11175
      Begin VB.CommandButton cmdMain 
         BackColor       =   &H000000C0&
         Caption         =   "Go to main page"
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
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5280
         Width           =   1815
      End
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats on the purchase of your new kitten!"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmKittenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'main kitten page, shows user the kitten they have purchased


Private Sub cmdMain_Click()
frmKittenForm.Hide
Welcomeform2.Show
End Sub

Private Sub lblCongrats_Click()


End Sub
