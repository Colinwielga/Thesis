VERSION 5.00
Begin VB.Form frmNothing 
   Caption         =   "SJU win's again! (Project by: Dan Gregus)"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   Picture         =   "beating st thomas form.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Nothing!"
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frmNothing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmNothing (beating st thomas form.frm)
'Dan Gregus
'3/19/06
'Objective: A very simple page with a very simple concept.  Displays a full background picture with a simple button bringing the user back to the home page.


Private Sub cmdBack_Click()
    frmNothing.Visible = False
    frmSJULacrosse.Visible = True
End Sub
