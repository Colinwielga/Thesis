VERSION 5.00
Begin VB.Form Status 
   BackColor       =   &H80000001&
   Caption         =   "Ace Hardware!  Home of the Helpful Hardware Folks!"
   ClientHeight    =   8565
   ClientLeft      =   2490
   ClientTop       =   1065
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdForward 
      BackColor       =   &H000040C0&
      Caption         =   "Advance To Reports"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit Program"
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdvance 
      BackColor       =   &H000040C0&
      Caption         =   "Advance To POS"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton cmdStartInfo 
      BackColor       =   &H00C000C0&
      Caption         =   "Display Current Status"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   2535
   End
   Begin VB.PictureBox picStartInv 
      BackColor       =   &H80000005&
      Height          =   9615
      Left            =   13080
      ScaleHeight     =   9555
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picPrices 
      BackColor       =   &H80000005&
      Height          =   9615
      Left            =   9240
      ScaleHeight     =   9555
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.PictureBox picItems 
      BackColor       =   &H80000005&
      Height          =   9615
      Left            =   3600
      ScaleHeight     =   9555
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblStoreInfo 
      BackColor       =   &H0080FFFF&
      Caption         =   $"Form1.frx":0BCE
      Height          =   1935
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblStartInv 
      BackColor       =   &H0000FFFF&
      Caption         =   "Starting Inventory:"
      Height          =   255
      Left            =   13080
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblPrices 
      BackColor       =   &H0000FFFF&
      Caption         =   "Prices:"
      Height          =   255
      Left            =   10320
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblItems 
      BackColor       =   &H0000FFFF&
      Caption         =   "Items In Inventory:"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdvance_Click()
        'jump from status to pointofsale
        Status.Visible = False
        PointOfSale.Visible = True
End Sub

Private Sub cmdForward_Click()
        'this button jumps to inventory info
        Status.Visible = False
        InventoryInfo.Visible = True
        
End Sub

Private Sub cmdQuit_Click()
        End
End Sub

Private Sub cmdStartInfo_Click()
        'clear picture boxes and display a blank line before data
        picItems.Cls
        picPrices.Cls
        picStartInv.Cls
        picItems.Print ""
        picPrices.Print ""
        picStartInv.Print ""
        
        'Acquire Input from file
        Open App.Path & "\RetailPOSAndInventoryControl.txt" For Input As #1
        Do Until EOF(1)
            Pos = 0
            Input #1, SItems, SPrices, SStartInv
            Pos = Pos + 1
            Items(Pos) = SItems
            Prices(Pos) = SPrices
            StartInv(Pos) = SStartInv
        
        'Print Items, Prices, and Starting Quantities
            picItems.Print Tab(2); Items(Pos)
            picPrices.Print Tab(2); FormatCurrency(Prices(Pos))
            picStartInv.Print Tab(2); StartInv(Pos)
        Loop
        Close #1
End Sub



'RetailPOSandInventoryControl program; status form
'this code was written on Tuesday, October 31, 2006
'edited on Thursday, November 1, 2006
'written by Mark Collette
'the purpose of this form is to read input from a text file
'the input was then put into arrays, and displayed in pictureboxes
'the subroutines read input from the file and put it into the arrays, end the program, or jump between forms
