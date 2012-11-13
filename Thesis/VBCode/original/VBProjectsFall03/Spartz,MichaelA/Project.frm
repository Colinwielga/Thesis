VERSION 5.00
Begin VB.Form Cheapest 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Cheapest Restaurant"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox types 
      BackColor       =   &H8000000D&
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   7155
      TabIndex        =   6
      Top             =   240
      Width           =   7215
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00C0FFC0&
      Height          =   2295
      Left            =   4200
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00800080&
      Caption         =   "Close"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Display 
      BackColor       =   &H00C00000&
      Caption         =   "Least expensive to most expensive restaurants according to your wants"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Wants 
      BackColor       =   &H00C0C000&
      Caption         =   "What would you like?"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox results 
      BackColor       =   &H8000000D&
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   720
      Width           =   7215
   End
   Begin VB.CommandButton Showoriginal 
      BackColor       =   &H0000FF00&
      Caption         =   "Show restaurants in area"
      Height          =   855
      Left            =   360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Cheapest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cheapest Restaurant Program (Restaurants)
'Cheapest Restuant (Cheapest)

'Developed By: Michael Spartz
'Written: 10/30/2003
        'Purpose:
        'Using six restuarants, display the
        'restuarants from cheapest to most expensive
        'depending on what the user wants,found out
        'through inputboxes,(how many:pops,boxes of
        'fries, hamburgers, and cheeseburgers)from
        'the restaurant

Option Explicit
Dim PATH As String
Dim IMAGENAME(1 To 6) As String
Dim i As Integer
Dim restaurant(1 To 6) As String
Dim pop(1 To 6) As Double
Dim fries(1 To 6) As Double
Dim burger(1 To 6) As Double
Dim cheeseburger(1 To 6) As Double
Dim wantpop As Double
Dim wantfries As Double
Dim wantburger As Double
Dim wantcheese As Double
Dim totalpop(1 To 6) As Double
Dim totalfries(1 To 6) As Double
Dim totalburgers(1 To 6) As Double
Dim totalcheese(1 To 6) As Double
Dim subtotal(1 To 6) As Double

Private Sub Display_Click()
Dim pass As Integer
Dim temp As Double
Dim short As String
results.Cls
types.Cls
For i = 1 To 6
totalpop(i) = (wantpop) * pop(i)
totalfries(i) = (wantfries) * fries(i)
totalburgers(i) = (wantburger) * burger(i)
totalcheese(i) = (wantcheese) * cheeseburger(i)
subtotal(i) = totalpop(i) + totalfries(i) + totalburgers(i) + totalcheese(i)
Next i

For pass = 1 To 6
    For i = 1 To 6 - pass
        If subtotal(i) > subtotal(i + 1) Then
            temp = subtotal(i)
            subtotal(i) = subtotal(i + 1)
            subtotal(i + 1) = temp
            short = restaurant(i)
            restaurant(i) = restaurant(i + 1)
            restaurant(i + 1) = short
            temp = totalpop(i)
            totalpop(i) = totalpop(i + 1)
            totalpop(i + 1) = temp
            temp = totalfries(i)
            totalfries(i) = totalfries(i + 1)
            totalfries(i + 1) = temp
            temp = totalburgers(i)
            totalburgers(i) = totalburgers(i + 1)
            totalburgers(i + 1) = temp
            temp = totalcheese(i)
            totalcheese(i) = totalcheese(i + 1)
            totalcheese(i + 1) = temp
            temp = pop(i)
            pop(i) = pop(i + 1)
            pop(i + 1) = temp
            temp = fries(i)
            fries(i) = fries(i + 1)
            fries(i + 1) = temp
            temp = burger(i)
            burger(i) = burger(i + 1)
            burger(i + 1) = temp
            temp = cheeseburger(i)
            cheeseburger(i) = cheeseburger(i + 1)
            cheeseburger(i + 1) = temp
            short = IMAGENAME(i)
            IMAGENAME(i) = IMAGENAME(i + 1)
            IMAGENAME(i + 1) = short
        End If
    Next i
Next pass
types.Print "Name of"; Tab(20); "Total price"; Tab(33); "Total price of"; Tab(49); "Total price"; Tab(66); "Total price of"; Tab(86); "Final"
types.Print "restuarant:"; Tab(20); "of pops:"; Tab(33); "boxes of fries:"; Tab(49); "of hamburgers:"; Tab(66); "cheeseburgers:"; Tab(86); "total:"
For i = 1 To 6
    results.Print restaurant(i); Tab(20); FormatCurrency(totalpop(i)); Tab(33); FormatCurrency(totalfries(i)); Tab(49); FormatCurrency(totalburgers(i)); Tab(67); FormatCurrency(totalcheese(i)); Tab(85); FormatCurrency(subtotal(i))
Next i
Pic.Picture = LoadPicture(IMAGENAME(1))
MsgBox "The cheapest restuarant is " & restaurant(1), , "Cheapest"
End Sub

Private Sub Form_Load()
PATH = "M:\CS130\Spartz, Michael A.\"
End Sub

Private Sub Quit_Click()
End
End Sub

Private Sub Showoriginal_Click()
results.Cls
types.Cls
Pic.Cls
Open PATH & "Names&prices.txt" For Input As #1
For i = 1 To 6
    Input #1, restaurant(i), pop(i), fries(i), burger(i), cheeseburger(i), IMAGENAME(i)
Next i
Close #1
types.Print "Name of"; Tab(20); "Price of"; Tab(33); "Price of"; Tab(49); "Price of"; Tab(69); "Price of"
types.Print "restuarant:"; Tab(20); "pop:"; Tab(33); "box of fries:"; Tab(49); "hamburger:"; Tab(69); "cheeseburger:"
For i = 1 To 6
    results.Print restaurant(i); Tab(20); FormatCurrency(pop(i)); Tab(33); FormatCurrency(fries(i)); Tab(49); FormatCurrency(burger(i)); Tab(69); FormatCurrency(cheeseburger(i))
Next i
End Sub

Private Sub Wants_Click()
results.Cls
wantpop = InputBox("How many pops do you want?", "pop")
wantfries = InputBox("How many boxes of fries do you want?", "fries")
wantburger = InputBox("How many hamburgers do you want?", "Hamburgers")
wantcheese = InputBox("How many cheeseburgers do you want?", "Cheeseburgers")
End Sub
