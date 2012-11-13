VERSION 5.00
Begin VB.Form frmBio 
   BackColor       =   &H80000007&
   Caption         =   "Biography"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmBio.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Bio"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   7695
      Left            =   3600
      ScaleHeight     =   7635
      ScaleWidth      =   10995
      TabIndex        =   2
      Top             =   3000
      Width           =   11055
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "frmBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Planet of Jet Li
'Form Name: frmBio
'Author: Chakong Thao
'Date Written: Monday, Oct. 30th
'Form Objective: This form is focused only on one subject, and that
                'is the biography of Jet Li.  This form limits its
                'space only for that use.
                
Option Explicit

Private Sub cmdBack_Click() 'This button hides this form and brings user back to General page
    frmBio.Hide
    frmGeneral.Show
End Sub

Private Sub cmdDisplay_Click()  'This button will load an array and display it into the picture box
    Dim Bio As String
    Open App.Path & "\Bio.txt" For Input As #1
    picResults.Cls
    
    Input #1, Bio
    
    Close #1
    
    picResults.Print Bio
    
End Sub

Private Sub cmdMain_Click() 'This button brings user back to the very beginning page/form
    frmBio.Hide
    frmJetLi.Show
End Sub
