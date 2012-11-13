VERSION 5.00
Begin VB.Form frmmap 
   Caption         =   "Map Of Seats"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   Picture         =   "frmmap.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MaskColor       =   &H000000C0&
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmmap; Form caption: Map of Seating

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmmap gives the user a map to of reference when
'                   choosing a ticket
Option Explicit


Private Sub cmdback_Click()
    frmmap.Hide
    frmtickets.Show
End Sub
