VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FF8080&
   Caption         =   "Form9"
   ClientHeight    =   11010
   ClientLeft      =   510
   ClientTop       =   885
   ClientWidth     =   14640
   LinkTopic       =   "Form9"
   ScaleHeight     =   11010
   ScaleWidth      =   14640
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      TabIndex        =   0
      Top             =   9720
      Width           =   1695
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForm2_Click()
Form9.Hide
Form2.Show
End Sub
