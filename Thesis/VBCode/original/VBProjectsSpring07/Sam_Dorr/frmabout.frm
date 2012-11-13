VERSION 5.00
Begin VB.Form frmabout 
   Caption         =   "About the CWS"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   10875
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
      Left            =   7440
      MaskColor       =   &H000000C0&
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdget 
      Caption         =   "Get the Info!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmabout; Form caption: About

'Author: Sam Dorr

'Date written: March 25, 2006

' Form Objective: The objective of frmabout is to give a brief history and explaination
'                   for the College World Seires
Option Explicit
Private Sub cmdback_Click()
frmabout.Hide
frmhome.Show
End Sub

Private Sub cmdget_Click()
Dim about As String

Open App.Path & "\about.txt" For Input As #1 'open text file
Do Until EOF(1) 'read text file
    Input #1, about
Loop
Close #1

MsgBox about, , "About" 'display txtfile

End Sub
