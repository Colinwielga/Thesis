VERSION 5.00
Begin VB.Form frmCongrats 
   BackColor       =   &H00000000&
   Caption         =   "Congratulations"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Mathematica6"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2655
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   600
      ScaleHeight     =   2355
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   3000
      Width           =   9375
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The Total Amount of Money won is:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   9495
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Congratulations!!!  You have completed the game of Jeopardy!!! "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmCongrats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()

    'Show and hide the forms
    frmCongrats.Hide
    frmCopyright.Show
    
End Sub
