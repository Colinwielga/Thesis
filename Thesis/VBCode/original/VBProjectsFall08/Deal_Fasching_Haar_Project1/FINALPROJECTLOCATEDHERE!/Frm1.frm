VERSION 5.00
Begin VB.Form frmtitle 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   Picture         =   "Frm1.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   14370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdworkscited 
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdleave 
      Caption         =   "Leave East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdenter 
      Caption         =   "Enter East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmtitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: High School Musical
' Form name: Title Page
' Author: Laura Deal, Megan Haar, Kirsten Fasching
' Date Written: 10/28/08
'Objective: this form in the title page of our project.  It allows the user to enter
' East High, which will bring the user to the table of contents, or the user can also
' leave East High, which is the exit button.

Option Explicit

Private Sub cmdenter_Click()

frmtitle.Hide
Frmcharacter.Hide
frmbuttons.Show
End Sub

Private Sub cmdleave_Click()
End
End Sub

Private Sub cmdworkscited_Click()
frmtitle.Hide
Frmcharacter.Hide
frmbuttons.Hide
frmworks.Show

End Sub