VERSION 5.00
Begin VB.Form MtnPro 
   Caption         =   "Bernese Mtn Dog"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14160
   LinkTopic       =   "Form3"
   ScaleHeight     =   11010
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<--Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton pr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next-->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   12225
      Left            =   -240
      Picture         =   "MtnPro.frx":0000
      Top             =   0
      Width           =   16350
   End
End
Attribute VB_Name = "MtnPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 If BackToPro < 2 Then
        MtnPro.Hide
        PupsPick.Show 'Goes back to profile after you have already clicked on back
    Else
        MtnPro.Hide
        ProMtn.Show
    End If
End Sub

Private Sub pr_Click()
If BackToPro < 2 Then 'Goes back to profile after you have already clicked on next
        MtnPro.Hide
        namePup.Show
    Else
        MtnPro.Hide
        ProMtn.Show
    End If
End Sub
