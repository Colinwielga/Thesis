VERSION 5.00
Begin VB.Form frmKGpic5 
   BackColor       =   &H8000000D&
   Caption         =   "KG, You Will Be Missed"
   ClientHeight    =   3600
   ClientLeft      =   5250
   ClientTop       =   3945
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5595
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
      Height          =   975
      Left            =   4080
      TabIndex        =   1
      Top             =   2280
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
      Height          =   975
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   0
      Picture         =   "frmKGpic5.frx":0000
      Top             =   0
      Width           =   4050
   End
End
Attribute VB_Name = "frmKGpic5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form displays a picture of KG. The two command buttons allow the user to return to the previous KG picture or return to the KG form

Private Sub cmdprevious_Click()
'This allows the user to view the previous KG picture
frmKGpic5.Visible = False
frmKGpic4.Visible = True

End Sub

Private Sub cmdreturn_Click()
'this allows the user to return to the KG main page
frmKGpic5.Visible = False
frmKG.Visible = True

End Sub
