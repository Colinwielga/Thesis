VERSION 5.00
Begin VB.Form frmBuilding 
   Caption         =   "Inside a building..."
   ClientHeight    =   11325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   Picture         =   "frmBuilding.frx":0000
   ScaleHeight     =   11325
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSign 
      Height          =   1455
      Left            =   5040
      Picture         =   "frmBuilding.frx":22C16
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSign_Click() 'show the sign with names
    frmBuilding.Hide
    frmSign.Show
    MsgBox ("It appears that the sign shows names of people...people who died and it tells their ages too.  This must be one place where aliens attacked fiercely."), , ("Names of the dead.")
End Sub
