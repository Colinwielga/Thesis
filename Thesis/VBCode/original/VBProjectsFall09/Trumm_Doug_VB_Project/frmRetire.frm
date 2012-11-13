VERSION 5.00
Begin VB.Form frmRetire 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Planning a retirement"
   ClientHeight    =   12255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   12255
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0FF&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   14040
      TabIndex        =   1
      Top             =   0
      Width           =   14100
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Invest earnings in stock market"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4200
      TabIndex        =   0
      Top             =   8640
      Width           =   5415
   End
End
Attribute VB_Name = "frmRetire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Use a random function to reckon how the user would fair investing their savings in the stock market.
    
Private Sub cmdRandom_Click()
    'Declare variables. Use Long to allow for very large values.
    Dim Portfolio As Long
    
    'Randomize in the range of 1 to 1,000,000
    Randomize
    Portfolio = Int(1000000 * Rnd) + 1
    
    'Use Select Case to sort randomized result into a category.
    Select Case Portfolio
        Case Is > 500000
            MsgBox ("You ride the blue-chip bubble into the ethanol euphoria and let petrol take you right on in the sunset. Your portfolio is worth " & FormatCurrency(Portfolio) & " you will live luxuriously in your retirement.")
            picResults.Picture = LoadPicture(App.Path & "\mrmoneybags.jpg")
        Case Is > 200000
            MsgBox ("You hit some hiccups along the way and have your bubble burst a few times, but are able to stumble into some halfway decent investments.  Your portfolio is valued at " & FormatCurrency(Portfolio) & " and you should be able to avoid some of the seedier retirement homes your ungrateful family may have tried to send you to.")
            picResults.Picture = LoadPicture(App.Path & "\abe.jpg")
        Case Is > 50000
            MsgBox ("Hope you aren't too proud to clip coupons.  You have only " & FormatCurrency(Portfolio) & " to live out the rest of your years.  Damn this infernal recession")
            picResults.Picture = LoadPicture(App.Path & "\coupons.jpg")
        Case Is >= 0
            MsgBox ("Pray with all your might that Social Security does not collapse on itself.  You squandered your money on poor investments.  Cherish the paltry " & FormatCurrency(Portfolio) & " that lingers pathetically in your account.")
            picResults.Picture = LoadPicture(App.Path & "\brokeguy.jpg")
    End Select
End Sub
