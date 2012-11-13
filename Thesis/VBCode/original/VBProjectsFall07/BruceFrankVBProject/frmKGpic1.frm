VERSION 5.00
Begin VB.Form frmKGpic1 
   Caption         =   "KG Slamming It Home"
   ClientHeight    =   9015
   ClientLeft      =   4005
   ClientTop       =   1680
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdnext 
      Caption         =   "See More KG Pics"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   1
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to KG page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   9390
      Left            =   0
      Picture         =   "frmKGpic1.frx":0000
      Top             =   0
      Width           =   8310
   End
End
Attribute VB_Name = "frmKGpic1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows a picture of KG and the two command buttons allow the user to either return to the KG form or go on to the next KG picture


Private Sub cmdNext_Click()
'Allows the user to proceed to the next picture of KG
frmKGpic1.Visible = False
frmKGpic2.Visible = True

End Sub

Private Sub cmdreturn_Click()
'Allows the user to return to the KG mainpage
frmKGpic1.Visible = False
frmKG.Visible = True

End Sub
