VERSION 5.00
Begin VB.Form frmPrice 
   BackColor       =   &H00FFFF80&
   Caption         =   "price"
   ClientHeight    =   7050
   ClientLeft      =   1905
   ClientTop       =   1635
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11070
   Begin VB.PictureBox picdog 
      Height          =   3495
      Left            =   6960
      Picture         =   "frmSize.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
   Begin VB.PictureBox picout 
      Height          =   3135
      Left            =   840
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox txtPrice 
      Height          =   1815
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Main Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6840
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter the maximum amount you wish to pay for a dog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dogs(VB-project.vbp)
'Form Name: frmSize (frmSize)
'Author: Libby Owen
'Date: Thursday Oct. 27
'Purpose: This form was written so that the user can input a price and
        'it will then give a list of the dogs less than or equal to that price
        


Private Sub cmdBack_Click()
' this button takes the user back to the previous page

frmPrice.Hide
frmFind.Show

End Sub

Private Sub cmdGo_Click()
' this button inputs the price the user put it.  It then compares the price to those in
'the array and will output prices less than or equal to that price

Dim kind(1 To 16) As String
Dim price(1 To 16) As Single
Dim I, x As Integer

x = txtPrice.Text

Open App.Path & "\Price.txt" For Input As #3
For I = 1 To 16
    Input #3, kind(I), price(I)
    If x > price(I) Then
        picout.Print kind(I); Tab(25); FormatCurrency(price(I)); Tab(60)
    End If
Next I

    

End Sub

