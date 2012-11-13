VERSION 5.00
Begin VB.Form Duchpro 
   Caption         =   "Duchshund"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form4"
   ScaleHeight     =   10155
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00004080&
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00004080&
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
      Height          =   1095
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "Duchpro.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14940
   End
End
Attribute VB_Name = "Duchpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If BackToPro < 2 Then 'Goes back to profile after you have already clicked on back
        Duchpro.Hide
        namePup.Show
    Else
        Duchpro.Hide
        Produch.Show
    End If
End Sub

Private Sub Command2_Click()
    If BackToPro < 2 Then 'Goes back to profile after you have already clicked on next
        Duchpro.Hide
        PupsPick.Show
    Else
        Duchpro.Hide
        Produch.Show
    End If

End Sub
