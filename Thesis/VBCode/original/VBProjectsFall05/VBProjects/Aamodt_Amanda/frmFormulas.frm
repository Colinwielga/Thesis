VERSION 5.00
Begin VB.Form frmFormulas 
   BackColor       =   &H00800080&
   Caption         =   "The Unit Circle, Some Identities, and Other Important Things"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   14985
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   12480
      TabIndex        =   7
      Text            =   "Designed by Amanda Aamodt"
      Top             =   9240
      Width           =   2415
   End
   Begin VB.PictureBox picTangent 
      Height          =   1335
      Left            =   11280
      Picture         =   "frmFormulas.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   3720
      Width           =   3495
   End
   Begin VB.PictureBox picTrig 
      BackColor       =   &H00C0FFC0&
      Height          =   4935
      Left            =   7200
      Picture         =   "frmFormulas.frx":1FE3E
      ScaleHeight     =   4875
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   480
      Width           =   3735
   End
   Begin VB.PictureBox picReciprocal 
      Height          =   1935
      Left            =   6840
      Picture         =   "frmFormulas.frx":64120
      ScaleHeight     =   1875
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   7560
      Width           =   5415
   End
   Begin VB.PictureBox picLawSines 
      Height          =   2295
      Left            =   360
      Picture         =   "frmFormulas.frx":932E2
      ScaleHeight     =   2235
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   7080
      Width           =   6135
   End
   Begin VB.PictureBox picLawCosines 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      Height          =   1695
      Left            =   8760
      Picture         =   "frmFormulas.frx":D5264
      ScaleHeight     =   1635
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   5640
      Width           =   5895
   End
   Begin VB.PictureBox picPyth 
      Height          =   3015
      Left            =   11160
      Picture         =   "frmFormulas.frx":FF16E
      ScaleHeight     =   2955
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.PictureBox picFormulas 
      BackColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   120
      Picture         =   "frmFormulas.frx":131190
      ScaleHeight     =   6555
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUnitCircle_Click()
    picFormulas.Cls
    picFormulas.Print
End Sub
