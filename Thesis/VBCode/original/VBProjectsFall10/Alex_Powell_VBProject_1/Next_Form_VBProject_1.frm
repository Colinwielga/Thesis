VERSION 5.00
Begin VB.Form frmShopping 
   BackColor       =   &H000000FF&
   Caption         =   "Shopping Cart"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   FillColor       =   &H8000000E&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10920
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForgotItems 
      BackColor       =   &H00000000&
      Caption         =   "Forgot Items"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H000000FF&
      TabIndex        =   8
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnterItems 
      BackColor       =   &H00000000&
      Caption         =   "Enter Items"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton cmdItemShow 
      BackColor       =   &H00000000&
      Caption         =   "Show Items"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtMyMerchandise 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2040
      Width           =   5175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrintTotal 
      BackColor       =   &H00000000&
      Caption         =   "Print Total"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   6480
      Width           =   2415
   End
   Begin VB.PictureBox picStoreName 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1080
      ScaleHeight     =   915
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.PictureBox picReceipt 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5295
      Left            =   6840
      ScaleHeight     =   5235
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblMerchandise 
      BackColor       =   &H000000FF&
      Caption         =   "Available Merchandise"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "frmShopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnterItems_Click()
    
    runningTotal = 0
    
    picReceipt.Print "Your Items"
    thePrice = InputBox("Enter a price, or -999 to show that you are done shopping")

    Do While thePrice <> -999
        picReceipt.Print FormatCurrency(thePrice, 2)
    
        runningTotal = runningTotal + thePrice
        
        thePrice = InputBox("Enter a price, or -999 to show that you are done shopping")
    Loop
    
    cmdEnterItems.Enabled = False
    
    
    
    
End Sub

Private Sub cmdForgotItems_Click()
'This button brings the same input box as before with the sentinel loop inside of it to enter data.
    'This is the presenting of the input box itself.
    thePrice = InputBox("Enter a price, or -999 to show that you are done shopping")
    'This is a loop statement that makes the input box stay there until all data is entered
    'along with a key stop(-999) to get rid of the input box.
    Do While thePrice <> -999
        'This prints each piece of data entered one by one into the picture box in the format of currency.
        picReceipt.Print FormatCurrency(thePrice, 2)
        'This keeps a continuous sum of the data coming from the input box.
        runningTotal = runningTotal + thePrice
        
        thePrice = InputBox("Enter a price, or -999 to show that you are done shopping")
    Loop
End Sub
Private Sub cmdItemShow_Click()
'This button prints a file into a text box that has a vertical scroll bar attatched to it.
    'Declaring new variables here
    Dim MyItem(1 To 500) As String
    Dim MyPrice(1 To 500) As Single
    Dim message As String, NextLine As String
    
    'This actually opens the file with the information needed to be printed into the text box.
    Open App.Path & "\MyItemsAndPrices.txt" For Input As #1
            
    message = "Here are the possible items to buy. "
        'This loops through all of the information so that it is read into an array.
        Do While Not EOF(1)
            'Increments the counter in the file.
            Ctr = Ctr + 1
            Input #1, MyItem(Ctr), MyPrice(Ctr)
            'This makes the message print one line at a time with the different information in the file.
            message = message & vbCrLf & MyItem(Ctr) & ",                        " & MyPrice(Ctr)
        Loop
        Close #1
        txtMyMerchandise.Text = message
End Sub
Private Sub cmdPrintTotal_Click()
'This button prints out the sum from the data above.
    picReceipt.Print Tab(18); "Your Total Is"
    picReceipt.Print "****************************************************************************"
    'This part makes it so that the sum is printed into the picture box with a currency format.
    picReceipt.Print Tab(21); FormatCurrency(runningTotal, 2)
End Sub

Private Sub cmdQuit_Click()
'This button ends the entire program.
End
End Sub




