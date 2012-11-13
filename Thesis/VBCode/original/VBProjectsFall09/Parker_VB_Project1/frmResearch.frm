VERSION 5.00
Begin VB.Form frmResearch 
   Caption         =   "Research"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   Picture         =   "frmResearch.frx":0000
   ScaleHeight     =   5445
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continue"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "List automakers by price"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlph 
      Caption         =   "List automakers alphabetically"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List automakers"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   3735
      Left            =   2880
      ScaleHeight     =   3675
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "frmResearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name is WickedFunCarBuilder
'Form Name is frmBegin
'Author is Dan Parker
'Date written 10/18/09
'The purpose of this form is to provide the user with the brand of a car, and give the user
'the average cost of every vehicle sold by the brand
Dim ctr As Integer

Private Sub cmdBack_Click()
'brings user back to homepage
    frmResearch.Hide
    frmFirst.Show
End Sub

Private Sub cmdCont_Click()
    'brings user to next page
    frmResearch.Hide
    frmGas.Show
End Sub

Private Sub cmdList_Click()
    Dim I As Integer 'dim local variable
    
    I = 0
    ctr = 0
    
    picResults.Print "Automaker"; Tab(25); Tab(35); "Average Cost Per Vehicle Sold" 'heading for picture box
    picResults.Print "**************************************************************************************************************"
    
    'load data into arrays
    Open App.Path & "\automakers.txt" For Input As #1

    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, autonames(ctr), autoprices(ctr)
    Loop
    Close #1
    
    'prints arrays
    For I = 1 To ctr
        picResults.Print autonames(I); Tab(25); Tab(45); FormatCurrency(autoprices(I))
    Next I
    
    Close #1
    picResults.Print " "
End Sub


Private Sub cmdQuit_Click()
MsgBox ("Thanks for using the Wicked Fun Car Builder, " & " " & UserName & "!")
End 'ends program
End Sub

Private Sub cmdAlph_Click()
    Dim Temp As String, Temp2 As Single, I As Integer, Pass As Integer, Pos As Integer 'dim local variables
    
    picResults.Cls 'clears picture box
    picResults.Print "Automaker"; Tab(25); Tab(35); "Average Cost Per Vehicle Sold"
    picResults.Print "**************************************************************************************************************"
    
    'sort function
    'arranges the names of the automakers in alphabetical order
    For Pass = 1 To ctr - 1
        For Pos = 1 To ctr - Pass
            If autonames(Pos) > autonames(Pos + 1) Then
                Temp = autonames(Pos)
                autonames(Pos) = autonames(Pos + 1)
                autonames(Pos + 1) = Temp
                Temp2 = autoprices(Pos)
                autoprices(Pos) = autoprices(Pos + 1)
                autoprices(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    'prints adjustted array
    For I = 1 To ctr
        picResults.Print autonames(I); Tab(25); Tab(45); FormatCurrency(autoprices(I), 2)
    Next I
End Sub

Private Sub cmdPrice_Click()
    Dim Temp As Single, Temp2 As String, I As Integer, Pass As Integer, Pos As Integer
    
    picResults.Cls
    picResults.Print "Automaker"; Tab(25); Tab(35); "Average Cost Per Vehicle Sold"
    picResults.Print "**************************************************************************************************************"
    
    'sort function
    'arranges the prices from most to least expensive
    For Pass = 1 To ctr - 1
        For Pos = 1 To ctr - Pass
            If autoprices(Pos) < autoprices(Pos + 1) Then
                Temp = autoprices(Pos)
                autoprices(Pos) = autoprices(Pos + 1)
                autoprices(Pos + 1) = Temp
                Temp2 = autonames(Pos)
                autonames(Pos) = autonames(Pos + 1)
                autonames(Pos + 1) = Temp2
            End If
        Next Pos
    Next Pass
    For I = 1 To ctr
    'prints adjusted array
        picResults.Print autonames(I); Tab(25); Tab(45); FormatCurrency(autoprices(I), 2)
    Next I
End Sub

