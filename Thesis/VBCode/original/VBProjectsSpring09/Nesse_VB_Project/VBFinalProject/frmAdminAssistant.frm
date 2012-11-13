VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H008080FF&
   Caption         =   "Available Positions"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "For more information and application instructions, click HERE."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdminAssistant.frx":0000
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label lblAdminAssistant 
      BackStyle       =   0  'Transparent
      Caption         =   "POSITION: Administrative Assistant"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Day at the Capitol and MN Private College Information Tool
'   Form: Admin
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of this form is to provide a job posting within the MPCC, and a link to further information about it online at www.mnprivatecolleges.org.

Private Sub cmdBack_Click()                     'Sends user back to previous page "Employment".

frmAdmin.Hide
frmHome.Show

End Sub

Private Sub lblLink_Click()                         'Like many other pages in application, allows link to Internet Explorer.
'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.mnprivatecolleges.org/employment/mpcc.php"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing
End Sub
