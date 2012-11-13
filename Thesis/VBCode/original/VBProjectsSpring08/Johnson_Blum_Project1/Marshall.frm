VERSION 5.00
Begin VB.Form Marshall 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmployment 
      Caption         =   "Want to become a Schwan's Truck Driver???"
      Height          =   1335
      Left            =   1800
      TabIndex        =   12
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Minnesota Home Page"
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdSortDessertPrice 
      Caption         =   "Sort Desserts by Price"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1680
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortDessert 
      Caption         =   "Sort Desserts Alphabetically"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortAppetizersPrice 
      Caption         =   "Sort Appetizers by Price"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortAppetizers 
      Caption         =   "Sort Appetizers Alphabetically"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortMeatPrice 
      Caption         =   "Sort Meat by Price"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortMeat 
      Caption         =   "Sort Meat Alphabetically"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdDesserts 
      Caption         =   "Desserts"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdSidesVegetables 
      Caption         =   "Appetizers"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   5055
      Left            =   3600
      ScaleHeight     =   4995
      ScaleWidth      =   5235
      TabIndex        =   2
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton cmdMeats 
      Caption         =   "Meat Products"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   7800
      Picture         =   "Marshall.frx":0000
      Top             =   120
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   240
      Picture         =   "Marshall.frx":4566
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label lblMarshall 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Marshall, Minnesota"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Marshall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Minnesoooota
'Form Name: Marshall
'Author: Danielle Johnson and Tony Blum
'Date Written: March 26th 2008
'The purpose of this form is load a file into an array and sort it by price or alphabetical order.  This form also allows you to switch forms and fill out an application to work at Schwan's.
Option Explicit
'Declares all variables as global throughout the form
Dim Desserts(1 To 50) As String, DessertCost(1 To 50) As Single, MeatCTR As Integer, AppCTR As Integer, DrtCTR As Integer
Dim Meat(1 To 50) As String, MeatCost(1 To 50) As Single
Dim Appetizers(1 To 50) As String, AppetizerCost(1 To 50) As Single

Private Sub cmdback_Click()
'Hides current form and returns to the Minnesota home page
Marshall.Hide
Minnesota.Show
End Sub

Private Sub cmdDesserts_Click()
'Initializes DrtCTR
DrtCTR = 0
'Loads the file "SchwansDeserts.txt" into two arrays-Desserts & DessertCost
Open App.Path & "\SchwansDeserts.txt" For Input As #1

Do While Not EOF(1)
    DrtCTR = DrtCTR + 1
    Input #1, Desserts(DrtCTR), DessertCost(DrtCTR)
Loop
Close #1
'Prints the arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"
For DrtCTR = 1 To DrtCTR
    picResults.Print Desserts(DrtCTR), Tab(60); FormatCurrency(DessertCost(DrtCTR))
Next DrtCTR
'Enables the Dessert sort buttons and dissables the current button
cmdSortDessert.Enabled = True
cmdSortDessertPrice.Enabled = True
cmdDesserts.Enabled = False
End Sub

Private Sub cmdEmployment_Click()
'Hides current form and goes to the Employment form
Marshall.Hide
Employment.Show
End Sub

Private Sub cmdMeats_Click()
'Initializes MeatCTR
MeatCTR = 0
'Loads the file "SchwansMeat.txt" into two arrays-Meat & MeatCost
Open App.Path & "\SchwansMeat.txt" For Input As #1

Do While Not EOF(1)
    MeatCTR = MeatCTR + 1
    Input #1, Meat(MeatCTR), MeatCost(MeatCTR)
Loop
Close #1
'Prints the arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"
For MeatCTR = 1 To MeatCTR
    picResults.Print Meat(MeatCTR), Tab(60); FormatCurrency(MeatCost(MeatCTR))
Next MeatCTR
'Enables the Dessert sort buttons and dissables the current button
cmdSortMeat.Enabled = True
cmdSortMeatPrice.Enabled = True
cmdMeats.Enabled = False
    
End Sub

Private Sub cmdSidesVegetables_Click()
'Initializes AppCTR
AppCTR = 0
'Loads the file "SchwansSidesVegetables.txt" into two arrays-Appetizers & AppetizerCost
Open App.Path & "\SchwansSidesVegetables.txt" For Input As #1

Do While Not EOF(1)
    AppCTR = AppCTR + 1
    Input #1, Appetizers(AppCTR), AppetizerCost(AppCTR)
Loop
Close #1
'Prints the arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"
For AppCTR = 1 To AppCTR
    picResults.Print Appetizers(AppCTR), Tab(60); FormatCurrency(AppetizerCost(AppCTR))
Next AppCTR
'Enables the Dessert sort buttons and dissables the current button
cmdSortAppetizers.Enabled = True
cmdSortAppetizersPrice.Enabled = True
cmdSidesVegetables.Enabled = False
End Sub

Private Sub cmdSortAppetizers_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the appetizers alphabetically
For Pass = 1 To AppCTR - 1
    For Pos = 1 To AppCTR - Pass
        If Appetizers(Pos) > Appetizers(Pos + 1) Then
            TempName = Appetizers(Pos)
            Appetizers(Pos) = Appetizers(Pos + 1)
            Appetizers(Pos + 1) = TempName
            TempNumber = AppetizerCost(Pos)
            AppetizerCost(Pos) = AppetizerCost(Pos + 1)
            AppetizerCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass
'Prints the alphabetized arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"

For P = 2 To AppCTR
    picResults.Print Appetizers(P), Tab(60); FormatCurrency(AppetizerCost(P))
Next P

End Sub



Private Sub cmdSortAppetizersPrice_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the appetizers by price
For Pass = 1 To AppCTR - 1
    For Pos = 1 To AppCTR - Pass
        If AppetizerCost(Pos) > AppetizerCost(Pos + 1) Then
            TempName = Appetizers(Pos)
            Appetizers(Pos) = Appetizers(Pos + 1)
            Appetizers(Pos + 1) = TempName
            TempNumber = AppetizerCost(Pos)
            AppetizerCost(Pos) = AppetizerCost(Pos + 1)
            AppetizerCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass
'Prints the arrays by cost with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"

For P = 2 To AppCTR
    picResults.Print Appetizers(P), Tab(60); FormatCurrency(AppetizerCost(P))
    
Next P
End Sub

Private Sub cmdSortDessert_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the desserts alphabetically
For Pass = 1 To DrtCTR - 1
    For Pos = 1 To DrtCTR - Pass
        If Desserts(Pos) > Desserts(Pos + 1) Then
            TempName = Desserts(Pos)
            Desserts(Pos) = Desserts(Pos + 1)
            Desserts(Pos + 1) = TempName
            TempNumber = DessertCost(Pos)
            DessertCost(Pos) = DessertCost(Pos + 1)
            DessertCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass
'Prints the alphabetized arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"

For P = 2 To DrtCTR
    picResults.Print Desserts(P), Tab(60); FormatCurrency(DessertCost(P))
Next P


End Sub

Private Sub cmdSortDessertPrice_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the desserts by price
For Pass = 1 To DrtCTR - 1
    For Pos = 1 To DrtCTR - Pass
        If DessertCost(Pos) > DessertCost(Pos + 1) Then
            TempName = Desserts(Pos)
            Desserts(Pos) = Desserts(Pos + 1)
            Desserts(Pos + 1) = TempName
            TempNumber = DessertCost(Pos)
            DessertCost(Pos) = DessertCost(Pos + 1)
            DessertCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass
'Prints the arrays by cost with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"

For P = 2 To DrtCTR
    picResults.Print Desserts(P), Tab(60); FormatCurrency(DessertCost(P))
Next P

End Sub

Private Sub cmdSortMeat_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the meats alphabetically
For Pass = 1 To MeatCTR - 1
    For Pos = 1 To MeatCTR - Pass
        If Meat(Pos) > Meat(Pos + 1) Then
            TempName = Meat(Pos)
            Meat(Pos) = Meat(Pos + 1)
            Meat(Pos + 1) = TempName
            TempNumber = MeatCost(Pos)
            MeatCost(Pos) = MeatCost(Pos + 1)
            MeatCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass
'Prints the alphabetized arrays with the appropriate header
picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"

For P = 2 To MeatCTR
    picResults.Print Meat(P), Tab(60); FormatCurrency(MeatCost(P))
Next P

        
End Sub

Private Sub cmdSortMeatPrice_Click()
Dim Pos As Integer, Pass As Integer, TempName As String, TempNumber As Single, P As Integer
'Sorts the desserts by price
For Pass = 1 To MeatCTR - 1
    For Pos = 1 To MeatCTR - Pass
    If MeatCost(Pos) > MeatCost(Pos + 1) Then
            TempName = Meat(Pos)
            Meat(Pos) = Meat(Pos + 1)
            Meat(Pos + 1) = TempName
            TempNumber = MeatCost(Pos)
            MeatCost(Pos) = MeatCost(Pos + 1)
            MeatCost(Pos + 1) = TempNumber
        End If
    Next Pos
Next Pass

picResults.Cls
picResults.Print "Product", Tab(60); "Price"
picResults.Print "--------------------------------------------------------------------------------------------------------------------------"
'Prints the arrays by cost with the appropriate header
For P = 2 To MeatCTR
    picResults.Print Meat(P), Tab(60); FormatCurrency(MeatCost(P))
Next P

End Sub
