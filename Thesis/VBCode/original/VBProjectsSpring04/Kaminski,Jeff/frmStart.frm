VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H0080C0FF&
      Caption         =   "Calculate My Grade "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "J. K., INC.    Grade Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()

frmStart.Hide
frmPctOne.Show

End Sub
