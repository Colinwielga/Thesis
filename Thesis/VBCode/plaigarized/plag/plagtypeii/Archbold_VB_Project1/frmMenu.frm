VERSION 5.00
Begin VB.Form frmMenu
   BackColor       =   &H00C0E0FF&
   Caption         =   "Menu"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShop
      Caption         =   "Start Shopping"
      Height          =   1455
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox thisPhotoOne
      Height          =   5655
      Left            =   9720
      ScaleHeight     =   5595
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   480
      Width           =   5055
   End
   Begin VB.PictureBox PicResults
      Height          =   5655
      Left            =   4200
      ScaleHeight     =   5595
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton tehalfy
      Caption         =   "Menu by Price"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton commandmoney
      Caption         =   "Menu alphabetically"
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name Sexton dining
'Form Name Menu
'Author Nick Archbold
'Date written 2/24/10
'Objective To plan out what to buy at sexton before punching out at the end of a week

Private Sub tehalfy_Click()
'sorts the items by ascending price, and calculates the total of all the items
    thisPhotoOne.Cls
Dim eye As Integer, pword As Integer, thelocation As Integer, notpermthing As String, permanentNot As Single, all As Integer
For pword = 1 To countthis - 1
     For thelocation = 1 To countthis - pword
        If Price(thelocation) > Price(thelocation + 1) Then
            notpermthing = item(thelocation)
            item(thelocation) = item(thelocation + 1)
            item(thelocation + 1) = notpermthing
                permanentNot = Price(thelocation)
                Price(thelocation) = Price(thelocation + 1)
                Price(thelocation + 1) = permanentNot
        End If
    Next thelocation
Next pword

For eye = 1 To countthis
    all = all + Price(eye)
    thisPhotoOne.Print item(eye), Tab(40), FormatCurrency(Price(eye))
Next eye
thisPhotoOne.Print
thisPhotoOne.Print "If you wanted to buy one of everything it would cost "; FormatCurrency(all)
End Sub

Private Sub commandmoney_Click()
'opens and reads file
Open App.Path & "\Menu.txt" For Input As #1
countthis = 0

Do While Not EOF(1)
    countthis = countthis + 1
    Input #1, item(countthis), Price(countthis)
Loop

Close #1
'dim variables
Dim eye As Integer, pword As Integer, thelocation As Integer, notpermthing As String, permanentNot As Single


PicResults.Cls
'sorts alphabetically
For pword = 1 To countthis - 1
     For thelocation = 1 To countthis - pword
        If item(thelocation) > item(thelocation + 1) Then
            notpermthing = item(thelocation)
            item(thelocation) = item(thelocation + 1)
            item(thelocation + 1) = notpermthing
                permanentNot = Price(thelocation)
                Price(thelocation) = Price(thelocation + 1)
                Price(thelocation + 1) = permanentNot
        End If
    Next thelocation
Next pword

For eye = 1 To countthis
    'prints list
    PicResults.Print item(eye), Tab(40), FormatCurrency(Price(eye))
Next eye
'disable alpahbet button
commandmoney.Enabled = False
'enables price button
tehalfy.Enabled = True

End Sub

Private Sub buythings_Click()
'goes to the store front
frmMenu.Visible = False
buythings.Visible = True
'asks how many punches the user wants to use and calculates the money value
currencies = InputBox("Please enter the number of punches you would like to use a punch is worth $4.85 at Sexton")
'display the monitary value
MsgBox (currencies & " punches is worth " & FormatCurrency((currencies * 4.85)))
End Sub

