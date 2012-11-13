VERSION 5.00
Begin VB.Form frmKGpic4 
   BackColor       =   &H8000000D&
   Caption         =   "Get That Stuff Out Rasheed"
   ClientHeight    =   6150
   ClientLeft      =   4830
   ClientTop       =   2100
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdNext 
      Caption         =   "See The Next KG Pic"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   4200
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
      Left            =   5040
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
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   0
      Picture         =   "frmKGpic4.frx":0000
      Top             =   0
      Width           =   4965
   End
End
Attribute VB_Name = "frmKGpic4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows a picture of KG.  The three command buttons allow the user to return to the KG page, return to the previous picture or proceed to the next picture.

Private Sub cmdNext_Click()
'This allows the user to proceed to the next KG picture
frmKGpic4.Visible = False
frmKGpic5.Visible = True

End Sub

Private Sub cmdprevious_Click()
'This allows the user to view the previous KG picture
frmKGpic4.Visible = False
frmKGpic3.Visible = True

End Sub

Private Sub cmdreturn_Click()
'This allows the user to return to the KG main page
frmKGpic4.Visible = False
frmKG.Visible = True

End Sub
