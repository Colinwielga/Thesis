VERSION 5.00
Begin VB.Form frm3 
   BackColor       =   &H000040C0&
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      Caption         =   "End Program"
      Height          =   975
      Left            =   10320
      TabIndex        =   3
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton cmdstats3 
      Caption         =   "Click to view stats"
      Height          =   975
      Left            =   5640
      TabIndex        =   2
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton Cmdreturn2 
      Caption         =   "Click to return to main page"
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   8640
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   480
      Picture         =   "frm3.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   12915
      TabIndex        =   0
      Top             =   1200
      Width           =   12975
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "Created By: Kyle Hinners"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "        St. John's Lacrosse Team Picture"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'project1 (project1.vbp)
'frm3(form3.frm)
'Kyle Hinners
'03/13/04
'The purpose of this form is to display the team picture, but also allows the user to go to the main page or the stats page



Private Sub cmdend_Click()
'this ends the program
End
End Sub






Private Sub Cmdreturn2_Click()
'this shows form 1 and hides form 3
frm1.Show
frm3.Hide
End Sub

Private Sub cmdstats3_Click()
'this shows form 2 and hides form 3
frm2.Show
frm3.Hide
End Sub
