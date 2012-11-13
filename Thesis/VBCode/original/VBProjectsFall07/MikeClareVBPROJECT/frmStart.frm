VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H0000C000&
   Caption         =   "ALIEN INVASION"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H0000FF00&
      Caption         =   "Play The Game!"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      MaskColor       =   &H0000FF00&
      TabIndex        =   5
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HELPFUL HINTS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALIEN INVASION!"
      BeginProperty Font 
         Name            =   "Lithos Pro Regular"
         Size            =   48
         Charset         =   0
         Weight          =   850
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   11055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStart.frx":0000
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2175
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You think you can survive an alien attack?  Let's find out"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H.P. = Health Points, this is how healthy you are.  0 H.P. means you're dead"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   5400
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Money = How much cash you have.  You may be able to buy helpful items."
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5760
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack = How strong you are.  If you have weapons, you're attack will rise.  You need to be strong to survive an alien fight!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   6120
      Width           =   6375
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdPlay_Click() 'when clicked the game begins.
    frmTitle.Hide
    frmStart.Show
End Sub
  
