VERSION 5.00
Begin VB.Form Checkout 
   BackColor       =   &H00000040&
   Caption         =   "Form2"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   5430
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Cmdshipping 
      Caption         =   "Compute Shipping"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton stotal 
      Caption         =   "Show Total"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox results 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   3000
      ScaleHeight     =   5055
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Checkout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : BodyKitOrder (Mike Seifert VB Project.vbp)
'Form Name : Checkout (Form2.frm)
'Author: Mike Seifert
'Date Written: November 5, 2003
'Purpose of Form: To allow the user to see and select from various
        'body kit styles for purchase, and then to tally the total
        'cost of the purchase and compute shipping.
        
Private Sub Form_Load()

Dim path As String
path = "App.path"

End Sub

Private Sub stotal_Click()


'open the file containing prices for various kit elements for use in an array
Open path & "prices.txt" For Input As #1

'load the array
For k = 1 To 6
    Input #1, prices(k)
Next k

'close the input file
Close #1

subtotal = 0

'print selected items for purchase and their quantity
results.Print "Items you have chosen to purchase today."
results.Print "-------------------------------------------"
results.Print "Part Name", , , "Price", "Quantity"

If aasbump > 0 Then
    results.Print "Andy's Autosport Front Bumper", FormatCurrency(prices(1), 2), aasbump
End If

If aaskirt > 0 Then
    results.Print "Andy's Autosport Side Skirts", , FormatCurrency(prices(2), 2), aaskirt
End If

If shogubump > 0 Then
    results.Print "Eribuni Shogun Front Bumper", FormatCurrency(prices(3), 2), shogubump
End If

If soskirt > 0 Then
    results.Print "Eribuni Shogun Side Skirts", , FormatCurrency(prices(4), 2), soskirt
End If

If budbump > 0 Then
    results.Print "Buddy Club Front Bumper", , FormatCurrency(prices(5), 2), budbump
End If

If buddyskirt > 0 Then
    results.Print "Buddy Club Side Skirts", , FormatCurrency(prices(6), 2), buddyskirt
End If
results.Print "------------------------------------------------"

'calculate the overall subtotal
subtotal = subtotal + aasbump * prices(1) + aaskirt * prices(2) + shogubump * prices(3) + soskirt * prices(4) + budbump * prices(5) + buddyskirt * prices(6)

'print the subtotal
results.Print "Your subtotal is", , FormatCurrency(subtotal, 2)

'compute tax
tax = 0.07 * subtotal
total = subtotal + tax

'print the total cost (subtotal + tax)
results.Print "Total cost is", , , FormatCurrency(total, 2)
results.Print ""
results.Print "If you are having the item(s) shipped, hit the 'compute shipping' button now."
End Sub

Private Sub Cmdshipping_Click()

'User enters shipping distance via an input box
distance = InputBox("Enter the distance in miles from 1666 S. Main St. Ste. A, Malpitas, CA 95035.  If you don't know the distance, you can find it using www.mapquest.com", "Find Distance")

'conditional statement to define shipping cost based on distance.
If distance < 100 Then
    shipping = 20
ElseIf distance < 200 Then
    shipping = 30
ElseIf distance < 300 Then
    shipping = 40
ElseIf distance < 400 Then
    shipping = 50
ElseIf distance < 500 Then
    shipping = 60
ElseIf distance < 600 Then
    shipping = 70
ElseIf distance < 700 Then
    shipping = 80
ElseIf distance < 800 Then
    shipping = 90
ElseIf distance > 800 Then
    shipping = 100
End If

'spacer.
results.Print ""

'prints the total shipping charge.
results.Print "Your shipping total will be "; FormatCurrency(shipping, 2)

'prints the total cost plus shipping.
results.Print "Total cost with shipping comes to "; FormatCurrency(total + shipping, 2)

End Sub
Private Sub CmdQuit_Click()
End
End Sub
