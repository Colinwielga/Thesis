VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   "About the Caclulator!"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHome 
      Caption         =   "Back to Home"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Image imgAbout 
      Height          =   6540
      Left            =   2640
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   9000
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright 2007 Willie Novitsky Design"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00000000&
      Caption         =   $"frmAbout.frx":EEFB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHome_Click()
    frmFirst.Show       'Returns to previous form
    frmAbout.Hide
End Sub

