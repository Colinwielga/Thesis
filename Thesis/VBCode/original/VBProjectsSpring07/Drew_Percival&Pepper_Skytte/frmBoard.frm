VERSION 5.00
Begin VB.Form frmBoard 
   Caption         =   "Your Board!"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   Picture         =   "frmBoard.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPrevious 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6120
      ScaleHeight     =   3915
      ScaleWidth      =   3315
      TabIndex        =   30
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton cmdNoDeal 
      BackColor       =   &H000000FF&
      Caption         =   "No Deal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8520
      Width           =   1935
   End
   Begin VB.PictureBox picOffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      ScaleHeight     =   1155
      ScaleWidth      =   3315
      TabIndex        =   27
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton cmdDeal 
      BackColor       =   &H0000FF00&
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8520
      Width           =   1935
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FF0000&
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   615
      Left            =   7080
      OleObjectBlob   =   "frmBoard.frx":1C206
      SourceDoc       =   "M:\CS130\Project 1-Deal or No Deal\Program Sounds\dond-bigthink.mp3"
      TabIndex        =   31
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Label lblYourBankOffer 
      BackColor       =   &H00000000&
      Caption         =   "The Banker has offered you...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   29
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "1000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   10080
      TabIndex        =   25
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "750000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   10080
      TabIndex        =   24
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "500000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   10080
      TabIndex        =   23
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "300000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   10080
      TabIndex        =   22
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "200000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   10080
      TabIndex        =   21
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "100000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   10080
      TabIndex        =   20
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "75000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   10080
      TabIndex        =   19
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "50000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   10080
      TabIndex        =   18
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "25000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   10080
      TabIndex        =   17
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   10080
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "5000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   10080
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   10080
      TabIndex        =   14
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "750"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   4080
      TabIndex        =   13
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   4080
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   4080
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4080
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4080
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   ".01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H0000FFFF&
      Caption         =   "400000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   10080
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form asks the user to pick deal or no deal and displays the board of values
'still left in cases

'The deal command button displays the offer that you have taken and the offer that
'was held in your case through a select case function
Private Sub cmdDeal_Click()

    'Using the CTRNoDeal value, the program determines what average to display
    Select Case CTRNoDeal
        Case 0
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average1, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))

        Case 1
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average2, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 2
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average3, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 3
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average4, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 4
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average5, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was "; FormatCurrency(Money(Pick))
            
        Case 5
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average6, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 6
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average7, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 7
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average8, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
            
        Case 8
            'Hides the board form
            frmBoard.Hide
            'Brings up the winner form
            frmWinner.Show
            'Displays what you won and what you would have won if you kept your case
            frmWinner.picValue.Print "Congratulations you have just "
            frmWinner.picValue.Print "won "; FormatCurrency(Average9, 0)
            frmWinner.picValue.Print "Your case that you sold "
            frmWinner.picValue.Print "was worth "; FormatCurrency(Money(Pick))
    
    'End the loop
    End Select
    
End Sub

'The NoDeal command button displays a message box telling the user how many cases to
'open. Then it hides the Board form and shows the DealOrNoDeal form
'If it is the ninth time the user selects the button then it displays the money in
'the users case on the winner form and shows the winner form
Private Sub cmdNoDeal_Click()

'Add one to CTRNoDeal every time that the NoDeal command button is used
CTRNoDeal = CTRNoDeal + 1

'Using the CTRNoDeal value, the program determines what round of the game you are in
Select Case CTRNoDeal
    
    Case 1
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 5 more cases.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 2
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 4 more cases.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 3
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 3 more cases.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 4
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 2 more cases.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 5
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 1 more case.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 6
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 1 more case.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 7
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 1 more case.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 8
        'Displays a message box that tells the user how many more cases to select
        MsgBox "Please select 1 more case.", , "Select"
        'Hides the board form
        frmBoard.Hide
        'Brings up the DealOrNoDeal form
        frmDealOrNoDeal.Show
        
    Case 9
        'Hides the Board form
        frmBoard.Hide
        'Brings up the Winner form
        frmWinner.Show
        'Displays the money value in you case that you have won
        frmWinner.picValue.Print "Congratulations you have"
        frmWinner.picValue.Print "decided to open your case."
        frmWinner.picValue.Print "Your case held "; FormatCurrency(Money(Pick))

'End the loop
End Select

'Clears the picture box that shows the offer
picOffer.Cls
'Clears the picture box that shows the previous offers
picPrevious.Cls

End Sub

