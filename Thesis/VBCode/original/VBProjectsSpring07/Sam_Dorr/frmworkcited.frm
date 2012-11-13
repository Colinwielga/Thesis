VERSION 5.00
Begin VB.Form frmworkcited 
   BackColor       =   &H000000C0&
   Caption         =   "Work Cited"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9930
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
      Left            =   3960
      MaskColor       =   &H000000C0&
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.maps.com"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   9255
   End
   Begin VB.Label lblworkcited2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.cwsomaha.com/html/home/index.asp"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9255
   End
   Begin VB.Label lblworkcited 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work Cited"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9255
   End
End
Attribute VB_Name = "frmworkcited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmworkcited; Form caption: Work Cited

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmsworkcited gives credit to pictures and information
'                   via text boxes.

Private Sub cmdback_Click()
    frmworkcited.Hide
    frmhome.Show
End Sub
