VERSION 5.00
Begin VB.Form frmpolice 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Police Station"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Run away"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lets see how drunk you are"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2775
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Class           =   "MPlayer"
      Height          =   375
      Left            =   4320
      OleObjectBlob   =   "frmpolice.frx":0000
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label lblpolice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Police Station"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   7260
      Left            =   -1920
      Picture         =   "frmpolice.frx":5018
      Top             =   -120
      Width           =   13470
   End
End
Attribute VB_Name = "frmpolice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
    'Project name:  Tour De St. Joe
    'Form:  frmpolice, "Police"
    'Author:  Brooke and Josh
    'Date:  3/115/08
    'Objective: Allows navigation between the main form and the police form

Private Sub cmdcalculate_Click()

    frmpolicebac.Show
    frmpolice.Hide

End Sub

Private Sub cmdleave_Click()
    
    frmpolicerun.Show
    frmpolice.Hide

End Sub

