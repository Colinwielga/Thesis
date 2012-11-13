VERSION 5.00
Begin VB.Form frmopening 
   Caption         =   "100 Meter Sprint"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   Picture         =   "frmopening.frx":0000
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00008000&
      Caption         =   "   Everything You Never Wanted To Know About The 100 Meter Sprint"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   10335
   End
End
Attribute VB_Name = "frmopening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdenter_Click()
    frmopening.Hide
    frmwhichfact.Show
End Sub
