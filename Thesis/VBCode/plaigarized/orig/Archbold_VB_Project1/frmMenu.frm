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
   Begin VB.PictureBox picresults1 
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
   Begin VB.CommandButton cmdAlphabet 
      Caption         =   "Menu by Price"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrice 
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

Private Sub cmdAlphabet_Click()
'sorts the items by ascending price, and calculates the total of all the items
    picresults1.Cls
Dim i As Integer, pass As Integer, pos As Integer, tempitem As String, tempPrice As Single, all As Integer
For pass = 1 To ctr - 1
     For pos = 1 To ctr - pass
        If Price(pos) > Price(pos + 1) Then
            tempitem = item(pos)
            item(pos) = item(pos + 1)
            item(pos + 1) = tempitem
                tempPrice = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = tempPrice
        End If
    Next pos
Next pass

For i = 1 To ctr
    all = all + Price(i)
    picresults1.Print item(i), Tab(40), FormatCurrency(Price(i))
Next i
picresults1.Print
picresults1.Print "If you wanted to buy one of everything it would cost "; FormatCurrency(all)
End Sub

Private Sub cmdPrice_Click()
'opens and reads file
Open App.Path & "\Menu.txt" For Input As #1
ctr = 0

Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, item(ctr), Price(ctr)
Loop

Close #1
'dim variables
Dim i As Integer, pass As Integer, pos As Integer, tempitem As String, tempPrice As Single


PicResults.Cls
'sorts alphabetically
For pass = 1 To ctr - 1
     For pos = 1 To ctr - pass
        If item(pos) > item(pos + 1) Then
            tempitem = item(pos)
            item(pos) = item(pos + 1)
            item(pos + 1) = tempitem
                tempPrice = Price(pos)
                Price(pos) = Price(pos + 1)
                Price(pos + 1) = tempPrice
        End If
    Next pos
Next pass

For i = 1 To ctr
    'prints list
    PicResults.Print item(i), Tab(40), FormatCurrency(Price(i))
Next i
'disable alpahbet button
cmdPrice.Enabled = False
'enables price button
cmdAlphabet.Enabled = True

End Sub

Private Sub cmdShop_Click()
'goes to the store front
frmMenu.Visible = False
frmStore.Visible = True
'asks how many punches the user wants to use and calculates the money value
Punches = InputBox("Please enter the number of punches you would like to use a punch is worth $4.85 at Sexton")
'display the monitary value
MsgBox (Punches & " punches is worth " & FormatCurrency((Punches * 4.85)))
End Sub

