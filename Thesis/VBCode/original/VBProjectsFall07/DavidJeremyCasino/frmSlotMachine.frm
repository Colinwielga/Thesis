VERSION 5.00
Begin VB.Form frmSlotMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slot Machine"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSlotMachine.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Slot Machine"
      Height          =   1215
      Left            =   1200
      Picture         =   "frmSlotMachine.frx":1B408
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSlotMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    frmSlotMachine.Hide
    frmCasino.Show
End Sub
