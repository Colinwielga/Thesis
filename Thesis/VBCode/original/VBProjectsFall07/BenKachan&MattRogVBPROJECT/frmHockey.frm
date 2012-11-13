VERSION 5.00
Begin VB.Form frmHockey 
   Caption         =   "Hockey Shop"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   Picture         =   "frmHockey.frx":0000
   ScaleHeight     =   7380
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort prices from highest to lowest"
      Height          =   1095
      Left            =   6600
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return Back to Main Menu!"
      Height          =   1095
      Left            =   6600
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdExpensive 
      Caption         =   "Sort list by prices more than $100.00"
      Height          =   1095
      Left            =   6600
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheap 
      Caption         =   "Sort list by prices less than $100.00"
      Height          =   1095
      Left            =   6600
      TabIndex        =   3
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "(MUST LOAD) Show me the items this store has to offer!!"
      Height          =   1095
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "WELCOME TO THE HOCKEY GIFT SHOP"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmHockey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Items(1 To 100) As String, Prices(1 To 100) As Single, CTR As Integer

Private Sub cmdCheap_Click()
'This subroutine searches the array for all items less than $100 and then displays them
Dim pass As Integer
    picResults.Cls
    For pass = 1 To CTR
        If (Prices(pass) < 100) Then
            picResults.Print Items(pass); Tab(40); FormatCurrency(Prices(pass))
        End If
    Next pass
End Sub

Private Sub cmdExpensive_Click()
    'This subroutine searches the array for all items greater than $100 and then displays them
    Dim pass As Integer
    picResults.Cls
    For pass = 1 To CTR
        If (Prices(pass) > 100) Then
             picResults.Print Items(pass); Tab(40); FormatCurrency(Prices(pass))
        End If
    Next pass
End Sub

Private Sub cmdLoad_Click()
    'This initially loads the data from a text file to an array and then displays the array.
    Dim pass As Integer
    picResults.Cls
    Open App.Path & "\Hockey.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Items(CTR), Prices(CTR)
    Loop
    Close #1
    picResults.Print "Items"; Tab(40); "Prices"
    picResults.Print "************************************************************"
    For pass = 1 To CTR
        picResults.Print Items(pass); Tab(40); FormatCurrency(Prices(pass))
    Next pass
End Sub

Private Sub cmdReturn_Click()
    'return to Main Menu
    frmHockey.Hide
    frmHome.Show
End Sub

Private Sub cmdSort_Click()
    'This subroutine bubble sorts the prices of items and also correctly moves the item name along with the shifted price
    Dim pass As Integer, Comp As Integer, tempItem As String, tempPrice As Single
    picResults.Cls
    For pass = 1 To CTR - 1
        For Comp = 1 To (CTR - pass)
            If (Prices(Comp) < Prices(Comp + 1)) Then
                tempItem = Items(Comp)
                Items(Comp) = Items(Comp + 1)
                Items(Comp + 1) = tempItem
                tempPrice = Prices(Comp)
                Prices(Comp) = Prices(Comp + 1)
                Prices(Comp + 1) = tempPrice
            End If
        Next Comp
    Next pass
    For pass = 1 To CTR
        picResults.Print Items(pass); Tab(40); FormatCurrency(Prices(pass))
    Next pass
End Sub
