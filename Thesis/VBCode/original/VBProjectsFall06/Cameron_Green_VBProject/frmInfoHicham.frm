VERSION 5.00
Begin VB.Form frmInfoHicham 
   BackColor       =   &H00008000&
   Caption         =   "Hicham El Guerrouj"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Runners Page"
      Height          =   855
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblHicham 
      BackColor       =   &H00008000&
      Caption         =   $"frmInfoHicham.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   840
      Picture         =   "frmInfoHicham.frx":022E
      Top             =   1560
      Width           =   2190
   End
   Begin VB.Label lblKing 
      BackColor       =   &H00008000&
      Caption         =   "The King of the Mile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmInfoHicham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to runner's information page, text on form shows accomplishments of the runner'
Private Sub cmdBack_Click()
    frmRunners.Show
    frmInfoHicham.Hide
End Sub
