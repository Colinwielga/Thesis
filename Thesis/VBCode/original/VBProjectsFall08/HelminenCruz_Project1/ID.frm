VERSION 5.00
Begin VB.Form frmidform 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10635
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ID.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoPetShop 
      BackColor       =   &H00FF8080&
      Caption         =   "Go find a pet"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      ScaleHeight     =   1275
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   4920
      Width           =   4695
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FF8080&
      Caption         =   "FIND ID"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   3735
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual PetWorld ID"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label lblEnterName 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter first name, middle initial and last name."
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   2895
   End
End
Attribute VB_Name = "frmidform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'creates an ID for the user before they can use the pet world


Private Sub cmdFind_Click()

Dim WholeName As String
Dim j As Integer
Dim first As String, middle As String, last As String
Dim id As String

picresults.Cls

WholeName = txtName.Text
j = InStr(WholeName, " ")
first = Left(WholeName, j - 1)
last = Right(WholeName, Len(WholeName) - (j + 2))
middle = Mid(WholeName, j + 1, 1)
id = Left(first, 1) & middle & Left(last, 6)
picresults.Print "Your ID is"; id


End Sub

Private Sub cmdGoPetShop_Click()

frmidform.Hide
Welcomeform2.Show




End Sub

Private Sub lblEnterName_Click()

End Sub
