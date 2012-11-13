VERSION 5.00
Begin VB.Form frmStreet2 
   Caption         =   "Back in the street..."
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   Picture         =   "frmStreet2.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue..."
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
      Left            =   4440
      TabIndex        =   0
      Top             =   6960
      Width           =   2055
   End
End
Attribute VB_Name = "frmStreet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdContinue_Click()  'enter into building
    frmStreet2.Hide
    frmBuilding.Show
    MsgBox ("Okay, now you head into a broken down building.  I wonder what you'll find"), , ("Into the building...")
    MsgBox ("You see a sign, click it to read what it says..."), , ("Click the sign")
End Sub
