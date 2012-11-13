VERSION 5.00
Begin VB.Form frmNationalCut 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton End 
      Caption         =   "Quit"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox picPicture 
      Height          =   2055
      Left            =   2400
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "CONGRATULATIONS!!!  Your score meets the NCAA Division III National Diving Championships minimum requirement!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmNationalCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Competitive Diving Form
'Form Name: frmNationalCut
'Marcus Rien
'3/22/09
'This loads a Picture and notifies the diver if they have qualified for the National Championships.
Private Sub Form_Load()
picPicture.Picture = LoadPicture(App.Path & "\DivingLogo.jpeg")
End Sub


