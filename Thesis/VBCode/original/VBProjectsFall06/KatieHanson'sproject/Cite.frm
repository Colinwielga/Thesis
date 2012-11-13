VERSION 5.00
Begin VB.Form Cite 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "Back to Title Page"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"Cite.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "Cite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Movies
'Form name: Cite
'Author: Katie Hanson
'Date Written: Nov 1 2006
'Objective: This form informs the user of where the information was from

Option Explicit
Private Sub cmdBack_Click()
    Title.Show
    Cite.Hide
End Sub
