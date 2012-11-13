VERSION 5.00
Begin VB.Form frmBianca 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   3180
   ClientTop       =   1485
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Dancer Page"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   0
      Picture         =   "frmBianca.frx":0000
      Top             =   -1800
      Width           =   8775
   End
End
Attribute VB_Name = "frmBianca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdreturn_Click()
'This command button allows me to return to my win a date with a dancer form
frmBianca.Visible = False
frmDancer.Visible = True

End Sub
