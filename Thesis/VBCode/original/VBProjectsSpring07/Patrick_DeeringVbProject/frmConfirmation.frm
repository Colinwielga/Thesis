VERSION 5.00
Begin VB.Form frmConfirmation 
   BackColor       =   &H00000000&
   Caption         =   "Confirmation Page"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Confirmation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.PictureBox picResult 
      Height          =   1335
      Left            =   2160
      ScaleHeight     =   1275
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   9120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblConfirmation 
      BackColor       =   &H00000000&
      Caption         =   "Confirmation Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmConfirmation.Hide
frmMilage.Show
'move to Milage form
End Sub

Private Sub cmdDone_Click()
If total > 250 Then
    MsgBox "Thank you for shopping with us " & fullname & "! Your item(s) will be delivered to " & location & " in approx. 4 - 6 hours.", , "Thank You!"
    Else
    MsgBox "Thank you for shopping with us " & fullname & "! Your item(s) will be delivered to " & location & " in approx. 2 - 4 hours.", , "Thank You!"
End If
'if statement was added to give company extra time to fill
'large orders (defined as >$250)
'programs inputs both name and dorm number into msgbox
End
End Sub



Private Sub cmdView_Click()
picResult.Print "Subtotal:"; Tab(20); FormatCurrency(subtotal)
Dim tax As Single

Select Case milage
    Case Is = 1
        picResult.Print "Delivery Fee:"; Tab(20); "Free!"
        tax = subtotal * 0.065
        picResult.Print "Tax:"; Tab(20); FormatCurrency(tax)
        total = subtotal + tax
        picResult.Print "Total:"; Tab(20); FormatCurrency(total)
    Case Is = 2
        picResult.Print "Delivery Fee:"; Tab(20); "$5.00"
        tax = (subtotal + 5) * 0.065
        picResult.Print "Tax:"; Tab(20); FormatCurrency(tax)
        total = subtotal + 5 + tax
        picResult.Print "Total:"; Tab(20); FormatCurrency(total)
    Case Is = 3
        picResult.Print "Delivery Fee:"; Tab(20); "$9.00"
        tax = (subtotal + 9) * 0.065
        picResult.Print "Tax:"; Tab(20); FormatCurrency(tax)
        total = subtotal + 9 + tax
        picResult.Print "Total:"; Tab(20); FormatCurrency(total)
    Case Is < 1
        MsgBox "Invalid delivery option selected"
    Case Is > 3
        MsgBox "Invalid delivery option selected"
    End Select
'case statements are set up to reflect where the user lives
'depending on their input, it will incorporate the correct fee.
    
    
End Sub
