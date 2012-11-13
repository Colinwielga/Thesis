VERSION 5.00
Begin VB.Form ShepProp 
   Caption         =   "Shepherd"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "<--Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Next-->"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   10005
      Left            =   0
      Picture         =   "ShepProp.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8745
   End
End
Attribute VB_Name = "ShepProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If BackToPro < 2 Then
        ShepProp.Hide
        namePup.Show 'Goes back to profile after you have already clicked on next
    Else
        ShepProp.Hide
        ProShep.Show
    End If
End Sub

Private Sub Command2_Click()
 If BackToPro < 2 Then
        ShepProp.Hide 'Goes back to profile after you have already clicked on next
        PupsPick.Show
    Else
        ShepProp.Hide
        ProShep.Show
    End If
End Sub
