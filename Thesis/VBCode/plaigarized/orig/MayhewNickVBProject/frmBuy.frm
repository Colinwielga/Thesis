VERSION 5.00
Begin VB.Form frmBuy 
   BackColor       =   &H0000FFFF&
   Caption         =   "Buy Stuff"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowRemainder 
      Caption         =   "Show me how much Cash and Stuff I've got left."
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturntoMain 
      Caption         =   "Return to Stall"
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdPitchers 
      Caption         =   "Buy Pitchers"
      Height          =   1215
      Left            =   8640
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdCups 
      Caption         =   "Buy Cups"
      Height          =   1215
      Left            =   8640
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdIce 
      Caption         =   "Buy Ice"
      Height          =   1215
      Left            =   5880
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdSugar 
      Caption         =   "Buy Sugar"
      Height          =   1215
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdLemons 
      Caption         =   "Buy Lemons"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin VB.PictureBox picCurrent 
      BeginProperty Font 
         Name            =   "@Gulim"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   1680
      Width           =   8295
   End
   Begin VB.Label lblBuyStuff 
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmBuy.frx":0000
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCups_Click()
    Dim Ccost As Single, tempCups As Integer
    Ccost = 5
    tempCups = InputBox("1 pack of cups (100 individual cups) cost " & FormatCurrency(Ccost) & ". How many packs would you like? Remember, you have " & FormatCurrency(Cash) & " remaining!")
    If tempCups < 0 Then
        MsgBox "Enter a positive number, duh."
    ElseIf tempCups * Ccost <= Cash Then
        Cups = Cups + tempCups * 100
        Cash = Cash - tempCups * Ccost
    Else
        MsgBox "Not enough cash!"
    End If
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub

Private Sub cmdIce_Click()
    Dim Icost As Single, tempIce As Integer
    Icost = 0.5
    tempIce = InputBox("Each tray of ice costs " & FormatCurrency(Icost) & ". How many trays would you like? Remember, you have " & FormatCurrency(Cash) & " remaining!")
    If tempIce < 0 Then
        MsgBox "Enter a positive number, duh."
    ElseIf tempIce * Icost <= Cash Then
        Ice = Ice + tempIce
        Cash = Cash - tempIce * Icost
    Else
        MsgBox "Not enough cash!"
    End If
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub

Private Sub cmdLemons_Click()
    Dim Lcost As Single, tempLemons As Integer
    Lcost = 1
    tempLemons = InputBox("Each lemon costs " & FormatCurrency(Lcost) & ". How many would you like? Remember, you have " & FormatCurrency(Cash) & " remaining!")
    If tempLemons < 0 Then
        MsgBox "Enter a positive number, duh."
    ElseIf tempLemons * Lcost <= Cash Then
        Lemons = Lemons + tempLemons
        Cash = Cash - tempLemons * Lcost
    Else
        MsgBox "Not enough cash!"
    End If
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub

Private Sub cmdPitchers_Click()
    Dim Pcost As Single, tempPitchers As Integer
    Pcost = 10
    tempPitchers = InputBox("Each pitcher costs " & FormatCurrency(Pcost) & ". How many would you like? Remember, you have " & FormatCurrency(Cash) & " remaining!")
    If tempPitchers < 0 Then
        MsgBox "Enter a positive number, duh."
    ElseIf tempPitchers * Pcost <= Cash Then
        Pitchers = Pitchers + tempPitchers
        Cash = Cash - tempPitchers * Pcost
    Else
        MsgBox "Not enough cash!"
    End If
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub

Private Sub cmdReturntoMain_Click()
    frmBuy.Hide
    frmMainScreen.Show
End Sub

Private Sub cmdShowRemainder_Click()
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub

Private Sub cmdSugar_Click()
    Dim Scost As Single, tempSugar As Integer
    Scost = 1.25
    tempSugar = InputBox("Each cup of sugar costs " & FormatCurrency(Scost) & ". How much would you like? Remember, you have " & FormatCurrency(Cash) & " remaining!")
    If tempSugar < 0 Then
        MsgBox "Enter a positive number, duh."
    ElseIf tempSugar * Scost <= Cash Then
        Sugar = Sugar + tempSugar
        Cash = Cash - tempSugar * Scost
    Else
        MsgBox "Not enough cash!"
    End If
    picCurrent.Cls
    picCurrent.Print "Cash remaining = "; FormatCurrency(Cash)
    picCurrent.Print "**********************************************************************************************************************************************************"
    picCurrent.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picCurrent.Print Lemons, Sugar, Ice, Cups, Pitchers
End Sub
