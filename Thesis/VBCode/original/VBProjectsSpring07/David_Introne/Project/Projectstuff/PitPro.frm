VERSION 5.00
Begin VB.Form PitPro 
   Caption         =   "PitPro"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
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
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
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
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<--Back"
      Height          =   615
      Left            =   8160
      TabIndex        =   1
      Top             =   11160
      Width           =   1815
   End
   Begin VB.CommandButton cmdPitPro 
      BackColor       =   &H8000000A&
      Caption         =   "Next-->"
      Height          =   855
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11760
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   11025
      Left            =   0
      Picture         =   "PitPro.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "PitPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
 If BackToPro < 2 Then
        PitPro.Hide
        PupsPick.Show 'Goes back to profile after you have already clicked on next
    Else
        PitPro.Hide
        Produch.Show
    End If
End Sub

Private Sub Command3_Click()
If BackToPro < 2 Then
        PitPro.Hide
        namePup.Show 'Goes back to profile after you have already clicked on next
    Else
        PitPro.Hide
        ProPit.Show
    End If
End Sub
