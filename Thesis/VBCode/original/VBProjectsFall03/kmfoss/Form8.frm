VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FF8080&
   Caption         =   "Form8"
   ClientHeight    =   11010
   ClientLeft      =   735
   ClientTop       =   885
   ClientWidth     =   14685
   LinkTopic       =   "Form8"
   ScaleHeight     =   11010
   ScaleWidth      =   14685
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
      Left            =   7200
      TabIndex        =   1
      Top             =   9720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Requirements for Application"
      BeginProperty Font 
         Name            =   "PaintStroke"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForm2_Click()
Form8.Hide
Form2.Show
End Sub
