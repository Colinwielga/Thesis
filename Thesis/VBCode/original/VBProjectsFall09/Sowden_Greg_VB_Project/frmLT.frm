VERSION 5.00
Begin VB.Form frmLT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Left Tackle"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBac16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to O-Line"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label lblLT 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmLT.frx":0000
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Football: The Offense
'   LT
'   Greg Sowden
'   10/18/09
'   this form toggles back to the OLine information form
 Private Sub cmdBac16_Click()
    frmLT.Hide
    frmOLine.Show
    
End Sub
