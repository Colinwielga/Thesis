VERSION 5.00
Begin VB.Form frmKGpic2 
   Caption         =   "KG stroking the J"
   ClientHeight    =   6750
   ClientLeft      =   5460
   ClientTop       =   1485
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdnext 
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
      Height          =   975
      Left            =   3840
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "See The Last Pic"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdreturn 
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
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmKGpic2.frx":0000
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "frmKGpic2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form shows a picture of KG.  The three command buttons allow the user to return to the KG page, return to the previous picture or proceed to the next picture.

Private Sub cmdNext_Click()
'This allows the user to proceed to the next KG picture
frmKGpic2.Visible = False
frmKGpic3.Visible = True

End Sub

Private Sub cmdprevious_Click()
'This allows the user to view the previous KG picture
frmKGpic2.Visible = False
frmKGpic1.Visible = True

End Sub

Private Sub cmdreturn_Click()
'This allows the user to return to the KG main page
frmKGpic2.Visible = False
frmKG.Visible = True

End Sub
