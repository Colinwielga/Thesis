VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H0000FFFF&
   Caption         =   "Find Breed"
   ClientHeight    =   8265
   ClientLeft      =   585
   ClientTop       =   1635
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picout 
      Height          =   7215
      Left            =   3960
      Picture         =   "frmFind.frx":0000
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   4
      Top             =   720
      Width           =   9615
      Begin VB.PictureBox picshow 
         Height          =   3615
         Left            =   600
         ScaleHeight     =   3555
         ScaleWidth      =   3195
         TabIndex        =   5
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdtemp 
      Caption         =   "Temperament"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdShed 
      Caption         =   "Shedding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to previous Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   5280
      Width           =   2655
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dogs(VB-project.vbp)
'Form Name: frmFind (frmFind.frm)
'Author: Libby Owen
'Date: Thursday Oct. 27
'Purpose: This form was written so that the user can put in specific input
        ' and be able to see some dogs that would work for their criteria
        



Option Explicit
Dim Breed(1 To 11) As String
Dim Amount(1 To 11) As String
Dim I, x As Integer


Private Sub cmdBack_Click()
frmFind.Hide
frmFirstscreen.Show

End Sub

Private Sub cmdPrice_Click()

' this button takes the user to a new form where they can input the price of the dog they are looking for

frmFind.Hide
frmPrice.Show


End Sub

Private Sub cmdShed_Click()
'This button shows the user a list of dogs and rate their shedding.  The info is from an array.
Open App.Path & "\shedding.txt" For Input As #1
For I = 1 To 11
    Input #1, Breed(I), Amount(I)
    picshow.Print Breed(I); Tab(25); Amount(I); Tab(60)
Next I
    


End Sub



Private Sub cmdtemp_Click()
'this button takes the user to the form where they can input what kind of temperament they are looking for

frmFind.Hide
frmTemp.Show

End Sub
