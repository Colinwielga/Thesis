VERSION 5.00
Begin VB.Form frmSorting 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go Back"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Click here to input your maximum price and the program will sort the dessert options for you!!!"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "frmSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'What's on the menu?
'frmSorting
'Michael Murakami ande Kevin Schnese
'30 October 2006
'This form allows you to sort the dessert menu based on a maximum amount you would like to pay.

Private Sub cmdBack_Click()
'This button allows the user to go between the regular menu and the dessert menu.
    frmMenu.Visible = False
    frmDessert.Visible = True
    frmSorting.Visible = False
End Sub
Private Sub cmdExit_Click()
'A message box will pop up after clicking the exit button and and thank the user for using our menu.
    MsgBox "Thank you!", , "THANK YOU!!!"
    End
End Sub
Private Sub cmdSort_Click()
'After the user has entered the maximum amount they wish to pay for dessert, the program will then output all desserts that fit into the users range. It the user enters a number outside of the specified range, a message box will pop up and tell the user that they have entered an invalid number.
    counter = 0
    picDisplay.Cls
    Open App.Path & "\dessert.txt" For Input As #1
        Do Until EOF(1)
            counter = counter + 1
            Input #1, dessertitems(counter), dessertprices(counter)
        Loop
    Close #1
    maximum = InputBox("Please enter in the maximum price amount in your price range ($2 to $7). No dollar sign ($) is necessary.", "Maximum Price Range")
    If maximum < 2 Then
        picDisplay.Cls
        picDisplay.Print "YOU HAVE ENTERED AN INVALID MAXIMUM PRICE!!"
    End If
    If maximum > 7 Then
        picDisplay.Cls
        picDisplay.Print "YOU HAVE ENTERED AN INVALID MAXIMUM PRICE!!"
    End If
    For Pass = 1 To 6
        For Pos = 1 To 6 - Pass
            If dessertprices(Pos) < dessertprices(Pos + 1) Then
                TempItem = dessertitems(Pos)
                dessertitems(Pos) = dessertitems(Pos + 1)
                dessertitems(Pos + 1) = TempItem
                TempPrice = dessertprices(Pos)
                dessertprices(Pos) = dessertprices(Pos + 1)
                dessertprices(Pos + 1) = TempPrice
            End If
        Next Pos
    Next Pass
    For Pos = 1 To 6
        If dessertprices(Pos) < maximum And maximum <= 7 Then
            picDisplay.Print dessertitems(Pos); ""; FormatCurrency(dessertprices(Pos))
        End If
    Next Pos
End Sub
