VERSION 5.00
Begin VB.Form frmweb 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Website"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox WebBrowser 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   480
      ScaleHeight     =   5595
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   1680
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "www.sanrio.com"
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open URL"
      Height          =   1335
      Left            =   4800
      Picture         =   "frmweb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H80000009&
      Caption         =   "Back to shop page"
      Height          =   1095
      Left            =   2880
      Picture         =   "frmweb.frx":4FF2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H80000009&
      Caption         =   "Back to the main page"
      Height          =   1095
      Left            =   4800
      Picture         =   "frmweb.frx":53FF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1695
   End
End
Attribute VB_Name = "frmweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
frmshop.Show
frmweb.Hide
End Sub

Private Sub cmdback_Click()
frmmain.Show
frmweb.Hide
End Sub


Private Sub cmdgo_Click()

         If Text1.Text <> "" Then
             WebBrowser.Navigate Text1.Text
             If WebBrowser1.Visible = False Then
                 WebBrowser1.Visible = True
             End If
         End If
         
End Sub


