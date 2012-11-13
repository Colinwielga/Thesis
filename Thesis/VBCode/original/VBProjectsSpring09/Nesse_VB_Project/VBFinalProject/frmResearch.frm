VERSION 5.00
Begin VB.Form frmResearch 
   BackColor       =   &H000000C0&
   Caption         =   "Research and Policymaking Sources"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdWeb 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please visit our website by clicking here to see more of our policymaking and research tools."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Day at the Capitol and MN Private College Information Tool
'   Form: Research
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of this form is to provide an easy tool for researchers and policymakers
'   to find more information, via MPCC's useful website. The user can also navigate home again.

Private Sub cmdWeb_Click()

'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://www.mnprivatecolleges.org/research/index.php"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing

End Sub

Private Sub Command2_Click()
frmResearch.Hide
frmAboutMPCC.Show

End Sub

