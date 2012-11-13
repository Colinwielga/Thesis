VERSION 5.00
Begin VB.Form graph 
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsales 
      Caption         =   "Compare Sales Per Share"
      Height          =   1095
      Left            =   6600
      TabIndex        =   5
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdbook 
      Caption         =   "Compare Book Prices"
      Height          =   1095
      Left            =   4560
      TabIndex        =   4
      Top             =   8520
      Width           =   1335
   End
   Begin VB.PictureBox picresults 
      Height          =   8055
      Left            =   240
      ScaleHeight     =   7995
      ScaleWidth      =   10875
      TabIndex        =   3
      Top             =   120
      Width           =   10935
   End
   Begin VB.CommandButton cmdearnings 
      Caption         =   "Compare Earnings Per Share"
      Height          =   1095
      Left            =   2520
      TabIndex        =   2
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdprice 
      Caption         =   "Compare Price"
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdswitch 
      Caption         =   "Switch to Main Form"
      Height          =   855
      Left            =   9240
      TabIndex        =   0
      Top             =   8760
      Width           =   1935
   End
End
Attribute VB_Name = "graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stocks
'written by Mat Eickhoff on 12 March 2004
Option Explicit
Dim path As String

Private Sub cmdbook_Click()
Set picresults.Picture = LoadPicture(path & "book.bmp", vbLPLarge, vbLPColor)
End Sub

Private Sub cmdearnings_Click()
Set picresults.Picture = LoadPicture(path & "earnings.bmp", vbLPLarge, vbLPColor)
End Sub

Private Sub cmdprice_Click()
picresults.Picture = LoadPicture(path & "price.bmp", vbLPLarge, vbLPColor)
End Sub

Private Sub cmdsales_Click()
Set picresults.Picture = LoadPicture(path & "sales.bmp", vbLPLarge, vbLPColor)
End Sub

Private Sub cmdswitch_Click()
graph.Hide
frmmain.Show
End Sub

Private Sub Form_Load()
path = "N:\CS130\handin\Eickhoff, Matthew\"
End Sub
