VERSION 5.00
Begin VB.Form frmChinese 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form4"
   ScaleHeight     =   7830
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBigBowl 
      BackColor       =   &H80000000&
      Caption         =   "Big Bowl"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblChinese 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   " Chinese Restaurants"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmChinese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBigBowl_Click()
frmBigBowl.Show
frmChinese.Hide
End Sub

Private Sub cmdChang_Click()
frmPFChang.Show
frmChinese.Hide
End Sub

