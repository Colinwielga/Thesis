VERSION 5.00
Begin VB.Form frmMap 
   Caption         =   "MN Capitol Complex Map"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   0
      Picture         =   "frmMap.frx":0000
      ScaleHeight     =   9075
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H000000C0&
         Caption         =   "Click to Go Back"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   8040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'   Day at the Capitol and MN Private College Information Tool
'   Form: Map
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of this form is to provide users with a map of the MN State Capitol grounds,
'   if they select the option to view it. The user can also navigate back to the DAC page.


Private Sub cmdBack_Click()
frmMap.Hide
frmDirections.Show
End Sub

