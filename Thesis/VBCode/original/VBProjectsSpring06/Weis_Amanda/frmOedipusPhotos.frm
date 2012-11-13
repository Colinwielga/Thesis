VERSION 5.00
Begin VB.Form frmOedipusPhotos 
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   1215
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblOedipus1 
      Caption         =   "OedipusTex, Billie Jo, and Greek Chorus"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   5880
      Width           =   3015
   End
End
Attribute VB_Name = "frmOedipusPhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGoBack_Click()
    frmOedipusPhotos.Hide
    frmOedipusCastList.Show
End Sub
