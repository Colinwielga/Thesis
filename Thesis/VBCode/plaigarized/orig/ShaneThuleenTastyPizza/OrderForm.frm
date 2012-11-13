VERSION 5.00
Begin VB.Form OrderForm 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWait 
      Caption         =   "While you Wait for your food"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   9360
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturnMenu 
      Caption         =   "Back to Menu"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   9360
      Width           =   2895
   End
   Begin VB.CommandButton cmdPizza 
      Caption         =   "Pizza"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton cmdDrinks 
      Caption         =   "Drinks"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   8400
      Width           =   2415
   End
   Begin VB.PictureBox picBill 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   720
      Width           =   7575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Start your order with Drinks"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label lblOrder 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Place your order here"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDrinks_Click()
    Dim Drink As String
    picBill.Print "Drinks"
    picBill.Print "***************************"
    'Load File
    Open App.Path & "/Drinks.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Drinks(CTR), Drinkcost(CTR)
    Loop
    Close #1
    'Search through the file to print order
    Do Until Drink = "done"
        Found = False
        Drink = InputBox("What would you like to drink?, when done ordering type done", "Drink Order")
        Pos = 0
        Do Until Found = True Or Pos >= CTR
        Pos = Pos + 1
        If LCase(Drinks(Pos)) = LCase(Drink) Or UCase(Drinks(Pos)) = UCase(Drink) Then
            Found = True
        End If
        Loop
        If Found = True Then
            picBill.Print Drinks(Pos), FormatCurrency(Drinkcost(Pos))
            DrinkTotal = DrinkTotal + Drinkcost(Pos)
        End If
        If Found = False Then
            MsgBox "Sorry, but we don't have that", vbOKOnly, "Sorry!"
        End If
    Loop
    'Print drink total
    picBill.Print "****************************"
    picBill.Print "Drink Total:", FormatCurrency(DrinkTotal)
    
End Sub

Private Sub cmdPizza_Click()
    Dim CTR As Integer, Pos As Integer, PizzaName As String, PizzaTotal As Single
    Dim Subtotal As Single, Total As Single, Found As Boolean, Tax As Single
    picBill.Print "****************************"
    picBill.Print "Pizza"
    picBill.Print "****************************"
    

    'Load the Pizza file
    Open App.Path & "/Pizza.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Pizza(CTR), Pcost(CTR)
    Loop
    Close #1
    'Search through the file to print order
    Do Until PizzaName = "done"
        Found = False
        PizzaName = InputBox("What kind of pizza would you like?, When you are done ordering type Done", "Drink Order")
        Pos = 0
        Do Until Found = True Or Pos >= CTR
        Pos = Pos + 1
        If LCase(Pizza(Pos)) = LCase(PizzaName) Or UCase(Pizza(Pos)) = UCase(PizzaName) Then
            Found = True
        End If
        Loop
        If Found = True Then
            picBill.Print Pizza(Pos), FormatCurrency(Pcost(Pos))
            PizzaTotal = PizzaTotal + Pcost(Pos)
        End If
    Loop
        If Found = False Then
            MsgBox "Sorry, but we don't have that", vbOKOnly, "Sorry!"
        End If
    'Print Drink Total
    picBill.Print "****************************"
    picBill.Print "Pizza Total:", FormatCurrency(PizzaTotal)
    picBill.Print "****************************"
    'Print total
    Subtotal = DrinkTotal + PizzaTotal
    Tax = Subtotal * 0.07
    Total = Subtotal + Tax
    picBill.Print "Subtotal", FormatCurrency(Subtotal)
    picBill.Print "Tax", "7%"
    picBill.Print "Total", FormatCurrency(Total)
    
End Sub



Private Sub cmdReturnMenu_Click()
    OrderForm.Visible = False
    Menu.Visible = True
End Sub

Private Sub cmdWait_Click()
    OrderForm.Visible = False
    MsgBox "Play Craps", vbOKOnly, "Dice"
    Dice.Visible = True
    
End Sub
