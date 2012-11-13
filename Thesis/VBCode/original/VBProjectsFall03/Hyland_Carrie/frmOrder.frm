VERSION 5.00
Begin VB.Form frmOrder 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Order"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9480
   FillColor       =   &H00FFFF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdMainMenu 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton cmdType 
      BackColor       =   &H00FF80FF&
      Caption         =   "Click to order!"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picKind 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3240
      ScaleHeight     =   1875
      ScaleWidth      =   4875
      TabIndex        =   12
      Top             =   2160
      Width           =   4935
   End
   Begin VB.CommandButton cmdFixings 
      BackColor       =   &H00FF8080&
      Caption         =   "First, click here to see all of the fixings to choose from!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit1 
      BackColor       =   &H00FF8080&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdReOrder 
      Caption         =   "Start Order Over!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      MaskColor       =   &H00FF00FF&
      TabIndex        =   1
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Designed by Carrie Hyland"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label lblVeggie 
      BackColor       =   &H00FFC0C0&
      Caption         =   "8.Vegetarian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblCarnitas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7.Carnitas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblBarbacoa 
      BackColor       =   &H00FFC0C0&
      Caption         =   "6.Barbacoa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblSteak 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5.Steak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblChicken 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4.Chicken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblBol 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3.Bol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTaco 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2.Taco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblBurrito 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1.Burrito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ProjChipotleOrder (Carrie Hyland's VB Project.vbp)
'Form Name: frmOrder (Order_form.frm)
'Author: Carrie Hyland
'Date Written: October 19, 2003
'Purpose of Form: To have a user order a Chipotle burrito/taco/bol with an choices of fixings and
                 ' add up subtotal and total.  To give the user adequate information about the fixings
                 ' that can be choosen by linking to the fixings form.
                 
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
'Declaring the variables that are global
Dim Burrito(1 To 5) As String, Price(1 To 5) As Single, Taco(1 To 5) As String
Dim PriceT(1 To 5) As Single, Bol(1 To 5) As String, PriceB(1 To 5) As Single


Private Sub cmdFixings_Click()
'Hiding the frmOrder and showing the frmFixings.  Making the
'button for Type visible.
frmOrder.Hide
frmFixings.Show
cmdType.Visible = True
End Sub

Private Sub cmdMainMenu_Click()
'Hide the Order form and show the Main Menu Form
frmOrder.Hide
frmMainMenu.Show
End Sub

Private Sub cmdQuit1_Click()
'End the program
End
End Sub

Private Sub cmdType_Click()
'Declaring the variables
Dim Types As Integer, Meat As String
Dim CTR As Integer, Total As Single, Guac As Single, EnterGuac As Integer
'Making the fixings button not visible.
cmdFixings.Visible = False
'Initializing the variables
Total = 0
Guac = 1.25
CTR = 0
'Opening the file for Burrito to be able to put it into an array
Open Path & "Burrito.txt" For Input As #1
'Taking the information from Burrito.txt and putting it
'into two arrays
Do While Not EOF(1)
   CTR = CTR + 1
   Input #1, Burrito(CTR), Price(CTR)
Loop
'Closing the file
Close

'Making an InputBox pop up to prompt the user for which type
'of product they would enjoy
Types = InputBox("Please indicate which type of product you would enjoy.  (Using 1 thru 3)", "Type")
'If the user wants a burrito, they type in a 1
If Types = 1 Then
    'Making an InputBox pop up to promp the user for what time
    'of meat they would enjoy
     Meat = InputBox("Please indicate which type of meat you would enjoy. (Using 4 thru 8)", "Meat")
     'Clear the picKind screen to make available for repeated use
     picKind.Cls
     'Print a title
     picKind.Print "Your Order is the following:"
      'If they want Chicken, they type a 4
      If Meat = 4 Then
          'Prints that they choose a chicken burrito and the price
          picKind.Print "You choose a "; Burrito(1); Tab(30); FormatCurrency(Price(1))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(1))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(1) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
              'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(1)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 5 Then
        'If the user wants a steak burrito
        'Prints that they choose a steak burrito and the price
          picKind.Print "You choose a "; Burrito(2); Tab(29); FormatCurrency(Price(2))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(2))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(2) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(2)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 6 Then
         'If the user wants a barbacoa burrito
        'Prints that they choose a barbacoa burrito and the price
          picKind.Print "You choose a "; Burrito(3); Tab(32); FormatCurrency(Price(3))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(3))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(3) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(3)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 7 Then
         'If the user wants a carnitas burrito
        'Prints that they choose a carnitas burrito and the price
          picKind.Print "You choose a "; Burrito(4); Tab(30); FormatCurrency(Price(4))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1. If you don't want any, type a 0. ")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(4))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(4) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(4)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 8 Then
         'If the user wants a vegetarian burrito
        'Prints that they choose a vegetarian burrito and the price
          picKind.Print "You choose a "; Burrito(5); Tab(33); FormatCurrency(Price(5))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1. If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(5))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(5) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(5)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        Else
        'If the user enters something besides 4-8, it prints the following:
          picKind.Print "Sorry, this is not an option, please pick again."
        End If
   ElseIf Types = 2 Then
   'The user chose tacos
   'Reinitializes the variable CTR to 0
   CTR = 0
   'Opens the file for Tacos to be able to be put into an array
   Open Path & "Taco.txt" For Input As #1
   'Starts the Do While Loop to put information into two arrays
   Do While Not EOF(1)
   CTR = CTR + 1
   Input #1, Taco(CTR), PriceT(CTR)
   Loop
   'Closes the file for Tacos
   Close
     'Making an inputbox appear to indicate the type of meat the user wants
     Meat = InputBox("Please indicate which type of meat you would enjoy. (Using 4 thru 8)")
     'Clears the picKind screen to allow for repeated use
     picKind.Cls
     'Prints a heading
     picKind.Print "Your order is the following:"
      If Meat = 4 Then
      'The user choose a chicken taco
          picKind.Print "You choose "; Taco(1); Tab(30); FormatCurrency(Price(1))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(1))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(1) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(1)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 5 Then
        'The user choose steak tacos
          picKind.Print "You choose "; Taco(2); Tab(29); FormatCurrency(Price(2))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(2))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =              $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(2) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(2)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 6 Then
        'The user chose barbacoa tacos
          picKind.Print "You choose "; Taco(3); Tab(28); FormatCurrency(Price(3))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1. If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(3))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(3) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(3)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 7 Then
        'The user chose carnitas tacos
          picKind.Print "You choose "; Taco(4); Tab(30); FormatCurrency(Price(4))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(4))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(4) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(4)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 8 Then
        'The user chose vegetarian tacos
          picKind.Print "You choose "; Taco(5); Tab(33); FormatCurrency(Price(5))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(5))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(5) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(5)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            End If
        Else
        'The user imputed something besides 4 thru 8
          picKind.Print "Sorry, this is not an option, please pick again."
       End If
  ElseIf Types = 3 Then
  'The user choose a bol
  'Reinitialize the variable CTR to 0
   CTR = 0
   'Open the file for Bol to be able to be put into an array
   Open Path & "Bol.txt" For Input As #1
   'Put information into two arrays
   Do While Not EOF(1)
   CTR = CTR + 1
   Input #1, Bol(CTR), PriceB(CTR)
   Loop
   'Closes the file
   Close
     'Making an inputbox appear to ask the user what type of meat they want
     Meat = InputBox("Please indicate which type of meat you would enjoy. (Using 4 thru 8)")
     'Clears the picKind screen to allow for repeated use
     picKind.Cls
     'Prints a heading
     picKind.Print "Your order is the following:"
      If Meat = 4 Then
      'The user chose a chicken bol
          picKind.Print "You choose "; Bol(1); Tab(25); FormatCurrency(Price(1))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1. If you don't want any, type a 0. ")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(1))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =              $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(1) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(1)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 5 Then
        'The use chose a steak bol
          picKind.Print "You choose "; Bol(2); Tab(23); FormatCurrency(Price(2))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(2))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =              $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(2) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(2)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 6 Then
        'The user chose a barbacoa bol
          picKind.Print "You choose "; Bol(3); Tab(28); FormatCurrency(Price(3))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1. If you don't want any, type a 0. ")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(3))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =               $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(3) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(3)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 7 Then
        'The user chose a carnitas bol
          picKind.Print "You choose "; Bol(4); Tab(25); FormatCurrency(Price(4))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(4))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =                $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(4) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(4)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        ElseIf Meat = 8 Then
        'The user chose a vegetarian bol
          picKind.Print "You choose "; Bol(5); Tab(28); FormatCurrency(Price(5))
          'Making an InputBox Pop up to decide if they want guac or not
          EnterGuac = InputBox("If you would like Guacamole for $1.25 extra, enter 1.  If you don't want any, type a 0.")
            'If they want Guacamole
            If EnterGuac = 1 Then
              'Print subtotal without guacamole
              picKind.Print "Your subtotal is:"; Tab(15); FormatCurrency(Price(5))
              'Add Guacamole and price to subtotal
              picKind.Print "Guac =               $1.25"
              'Prints a divider
              picKind.Print "----------------------------------"
              'Adds subtotal with guacamole to get total
              Total = Price(5) + Guac
              'Prints the total
              picKind.Print "Your Total is:", FormatCurrency(Total, 2)
            Else
            'If the user doesn't want guacamole, the total is simply
              'the price
              Total = Price(5)
              'Prints a divider
              picKind.Print "------------------------------------"
              'Prints the total
              picKind.Print "Your total is:", FormatCurrency(Total, 2)
            End If
        Else
        'The user chose something besides 4 thru 8
          picKind.Print "Sorry, this is not an option, please pick again."
       End If
        
  Else
  'The user chose something besides 1 to 3
    picKind.Print "Sorry, this is not an option, please pick again."
End If
'make the Fixings button invisible
cmdFixings.Visible = False
End Sub


