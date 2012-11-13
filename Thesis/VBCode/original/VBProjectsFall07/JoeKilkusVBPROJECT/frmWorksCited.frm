VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBacktoMain 
      Caption         =   "Back to the front page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewWorksCited 
      Caption         =   "View Works Cited"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.PictureBox picWorksCited 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   1200
      ScaleHeight     =   4995
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   2880
      Width           =   9495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   11655
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this button takes the user back to the main page
Private Sub cmdBacktoMain_Click()
    frmWorksCited.Hide
    frmSJU_CC.Show
End Sub

'pressing this button displays the works cited information
Private Sub cmdViewWorksCited_Click()
    picWorksCited.Print "All team information was taken from www.gojohnnies.com."
    picWorksCited.Print ""
    picWorksCited.Print "All pictures were taken by the parents of members of the SJU Cross Country Team."
    picWorksCited.Print ""
    picWorksCited.Print "All race results were taken from www.gojohnnies.com."
    picWorksCited.Print ""
    picWorksCited.Print "All code used for the program is out of the CSCI 130 text."
    
    cmdViewWorksCited.Enabled = False
    cmdBacktoMain.Enabled = True
    
End Sub
