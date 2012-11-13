VERSION 5.00
Begin VB.Form frmKGpic3 
   BackColor       =   &H8000000D&
   Caption         =   "KG Looking Cool"
   ClientHeight    =   5985
   ClientLeft      =   5250
   ClientTop       =   2100
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   5805
   Begin VB.CommandButton cmdNext 
      Caption         =   "See The Next Pic"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "See The Last KG Pic"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return To KG Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmKGpic3.frx":0000
      Top             =   0
      Width           =   4290
   End
End
Attribute VB_Name = "frmKGpic3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows a picture of KG.  The three command buttons allow the user to return to the KG page, return to the previous picture or proceed to the next picture.

Private Sub cmdNext_Click()
'This allows the user to proceed to the next KG picture
frmKGpic3.Visible = False
frmKGpic4.Visible = True

End Sub

Private Sub cmdprevious_Click()
'This allows the user to view the previous KG picture
frmKGpic3.Visible = False
frmKGpic2.Visible = True

End Sub

Private Sub cmdreturn_Click()
'This allwos the user to return to the KG main page
frmKGpic3.Visible = False
frmKG.Visible = True

End Sub
