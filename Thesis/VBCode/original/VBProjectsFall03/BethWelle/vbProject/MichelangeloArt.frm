VERSION 5.00
Begin VB.Form frmMichelangelo 
   BackColor       =   &H00C00000&
   Caption         =   "Michelangelo's Art"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to learn further information about Michelangelos's sculptures!"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   3840
      Width           =   5295
   End
   Begin VB.PictureBox pbxRondanini 
      Height          =   3495
      Left            =   7320
      Picture         =   "MichelangeloArt.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   4560
      Width           =   2655
   End
   Begin VB.PictureBox pbxPieta 
      Height          =   3495
      Left            =   3840
      Picture         =   "MichelangeloArt.frx":528A
      ScaleHeight     =   3435
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   4560
      Width           =   2775
   End
   Begin VB.PictureBox pbxMoses 
      Height          =   3495
      Left            =   360
      Picture         =   "MichelangeloArt.frx":F8D3
      ScaleHeight     =   3435
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.PictureBox pbxMadonna 
      Height          =   3375
      Left            =   7320
      Picture         =   "MichelangeloArt.frx":17DAA
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox pbxCrucifix 
      Height          =   3375
      Left            =   3840
      Picture         =   "MichelangeloArt.frx":1D5A5
      ScaleHeight     =   3315
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox pbxAtlas 
      Height          =   3375
      Left            =   360
      Picture         =   "MichelangeloArt.frx":21DD3
      ScaleHeight     =   3315
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   480
      ScaleHeight     =   135
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   240
      Width           =   15
   End
   Begin VB.Label lblArt 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click on a picture to learn it's name!  "
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
End
Attribute VB_Name = "frmMichelangelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Michelangelo Art
'Michelangelo Art (MichelangeloArt.frm)
'Beth Welle
'October 29, 2003
'Purpose of this form is to show user the different sculptures and name of each one.

Private Sub cmdSwitch_Click()
'switches from the first (picture) form to the second (informational) form.
frmMichelangelo.Hide
frmInfo.Show

End Sub

'Each of the pictures, when clicked, will state the name of the sculpture.

Private Sub pbxAtlas_Click()
MsgBox ("This is a sculpture of Atlas.")
End Sub

Private Sub pbxCrucifix_Click()
MsgBox ("This is a sculpture of the Crucifix.")
End Sub

Private Sub pbxMoses_Click()
MsgBox ("This is a sculpture of Moses.")
End Sub

Private Sub pbxPieta_Click()
MsgBox ("This is a sculpture of the Pieta.")
End Sub

Private Sub pbxRondanini_Click()
MsgBox ("This is a sculpture of the Rondanini.")
End Sub

Private Sub pbxMadonna_Click()
MsgBox ("This is a sculpture of Madonna with Child.")
End Sub
