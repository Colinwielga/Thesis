VERSION 5.00
Begin VB.Form FrmSlotMachine 
   Caption         =   "Big Money Slot Machine"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   Picture         =   "FrmSlotMachine.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicGoodLuck 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      ScaleHeight     =   1155
      ScaleWidth      =   8355
      TabIndex        =   15
      Top             =   960
      Width           =   8415
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CmdCashOut 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cash Out"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CmdChangeBet 
      BackColor       =   &H00FF00FF&
      Caption         =   "Change Bet"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CmdPlaceBet 
      BackColor       =   &H000080FF&
      Caption         =   "Place Bet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CmdSpin 
      BackColor       =   &H0000C000&
      Caption         =   "SPIN!!!!!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   5055
   End
   Begin VB.PictureBox Pic3 
      Height          =   3135
      Left            =   8520
      ScaleHeight     =   3075
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox Pic2 
      Height          =   3135
      Left            =   5760
      ScaleHeight     =   3075
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.PictureBox Pic1 
      Height          =   3135
      Left            =   2880
      ScaleHeight     =   3075
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.PictureBox PicWon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox PicMoney 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox PicBet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bets:             $0.10 $0.25 $0.50 Or Any Dollar Amount Up To $50.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   12120
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Won"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSlotMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Form is the main program, it allows the user to
'play slots.
'The background picture for this comes from:
'www.masgue.com

Option Explicit
Dim LegalBet(1 To 53) As Single 'LegelBet is used in both the initial placement of the bet and if the user wishes to change their bet during the game.
Dim Bet As Single   'Bet is used in three of the subroutines, PlaceBet, ChangeBet, and Spin.


'This subroutines is where the user makes their initial bet.
'The user must place a bet or the other buttons in the form
'will not work. The users starting money and intial winnings
'are set here as well.


Private Sub CmdPlaceBet_Click()
    Dim Found As Boolean
    Dim I As Integer
    Dim Ctr As Integer
    
    PicGoodLuck.Print "Good Luck "; UserName; "!!!!"
    Money = 100
    PicMoney.Print FormatCurrency(Money)
    Won = 0
    PicWon.Print FormatCurrency(Won)
    Bet = InputBox("Enter Your Bet", "Bet")
    Found = False
    I = 0
    Ctr = 0
    Open App.Path & "\LegalBet.txt" For Input As #1
    
    
    
    'This piece here verifies that the user has input a valid
    'amount. The valid amounts are stated on the side of the game screen.
    'If the user has entered a valid bet then the bet is stored on the
    'variables module, otherwise they are presented an error message
    'and forced to enter a new bet if they wish to continue.
    
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, LegalBet(Ctr)
    Loop
    
    Do While (Not Found) And (I < Ctr)
        I = I + 1
        If Bet = LegalBet(I) Then Found = True
    Loop
    
    If (Not Found) Then
        MsgBox "You Have Entered An Invalid Bet.", , "Error"
        PicWon.Cls
        PicMoney.Cls
        Else
            PicBet.Print FormatCurrency(Bet)
            CmdChangeBet.Enabled = True
            CmdCashOut.Enabled = True
            CmdSpin.Enabled = True
            CmdPlaceBet.Enabled = False
    End If
    
    Close #1
    
End Sub


'This subroutine allows the user to change the amount they are betting mid game.

Private Sub CmdChangeBet_Click()
    Dim NewBet As Single
    Dim Found As Boolean
    Dim I As Integer
    Dim Ctr As Integer
    
   
'Here the program is getting the new bet from an input box and then verifying the
'bet the same way as in the initial betting phase. Like the Initial phase if the
'bet is good, they continue with the game, if the bet is not valid they are again
'presented with a message box telling them their bet is not valid.  Their original
'bet remains until they put a valid new bet in.

    NewBet = InputBox("Enter Your New Bet", "Change Bet")
    Found = False
    I = 0
    Ctr = 0
    
    If NewBet > Money Then
        MsgBox "You Don't Have Enough Money For That Bet.", , "Error"
        Else
    
            Open App.Path & "\LegalBet.txt" For Input As #1
    
            Do While Not EOF(1)
                Ctr = Ctr + 1
                Input #1, LegalBet(Ctr)
            Loop
    
            Do While (Not Found) And (I < Ctr)
                I = I + 1
                If NewBet = LegalBet(I) Then Found = True
            Loop
    
            If (Not Found) Then
                MsgBox "You Have Entered An Invalid Bet.", , "Error"
                Else
                    Bet = NewBet
                    PicBet.Cls
                    PicBet.Print FormatCurrency(Bet)
            End If
    End If
    Close #1
End Sub

'This subroutine allows the user to quit while they are ahead.

Private Sub CmdCashOut_Click()
    MsgBox "Thanks For Playing!", , "Cash Out"
    End
End Sub

'If the user drops to zero dollars they are notiyied that they have no
'more money and are forced to quit.

Private Sub CmdQuit_Click()
    End
End Sub

'This is the Subroutine where the gambling takes place. The user's bet has been inputted and checked.
'The program takes three random numbers between one and four and Assigns them to variables. The the numbers
'one to four are each assigned a picture and when the program generates the three numbers each picture box displays
'the picture that has been assigned to the number. After that happens, an array is loaded containing the winning
'combinations, the program then searches for a match. If the program finds a match it then moves to assigning the
'odds to that winning combination. It then displays the amount won and changes the users total money. If the round was
'a loss then the amount bet is subtracted from the users total money.  This loop continues until the user decides to
'quit or runs out of money.
'The images for the slot machine come from:
'www.mesart.com
'www.apachenugget.com
'www.wholesaleapplique.com
'I found the code to generate random numbers at:
'http://msdn2.microsoft.com/en-us/library/8zedbtdt(VS.80).aspx

Private Sub CmdSpin_Click()
    Randomize
    
    Dim Number1 As Integer
    Dim Number2 As Integer
    Dim Number3 As Integer
    Dim Win As Boolean
    Dim Found As Boolean
    Dim Arr1(1 To 16) As Integer
    Dim Arr2(1 To 16) As Integer
    Dim Arr3(1 To 16) As Integer
    Dim Ctr As Integer
    Dim I As Integer
    
    Pic1.Cls
    Pic2.Cls
    Pic3.Cls
    
    
    Number1 = Int((4 * Rnd) + 1)
        Select Case Number1
            Case Is = 1
                Pic1.Picture = LoadPicture(App.Path & "\Cherry2.jpg")
            Case Is = 2
                Pic1.Picture = LoadPicture(App.Path & "\MoneyBag2.jpg")
            Case Is = 3
                Pic1.Picture = LoadPicture(App.Path & "\JackPot2.jpg")
            Case Is = 4
                Pic1.Picture = LoadPicture(App.Path & "\Bar2.jpg")
        End Select
        
    Number2 = Int((4 * Rnd) + 1)
        Select Case Number2
            Case Is = 1
                Pic2.Picture = LoadPicture(App.Path & "\Cherry2.jpg")
            Case Is = 2
                Pic2.Picture = LoadPicture(App.Path & "\MoneyBag2.jpg")
            Case Is = 3
                Pic2.Picture = LoadPicture(App.Path & "\JackPot2.jpg")
            Case Is = 4
                Pic2.Picture = LoadPicture(App.Path & "\Bar2.jpg")
        End Select
    
    Number3 = Int((4 * Rnd) + 1)
        Select Case Number3
            Case Is = 1
                Pic3.Picture = LoadPicture(App.Path & "\Cherry2.jpg")
            Case Is = 2
                Pic3.Picture = LoadPicture(App.Path & "\MoneyBag2.jpg")
            Case Is = 3
                Pic3.Picture = LoadPicture(App.Path & "\JackPot2.jpg")
            Case Is = 4
                Pic3.Picture = LoadPicture(App.Path & "\Bar2.jpg")
        End Select

     
    Open App.Path & "\WinningCombinations.txt" For Input As #1
        Ctr = 0
        Do While Not EOF(1)
            Ctr = Ctr + 1
            Input #1, Arr1(Ctr), Arr2(Ctr), Arr3(Ctr)
        Loop
        
    Win = False
    Found = False
    I = 0
    
    Do While (Not Found) And (I < Ctr)
        I = I + 1
        If Number1 = Arr1(I) And Number2 = Arr2(I) And Number3 = Arr3(I) Then Found = True
    Loop
    
    If Found = True Then
        Win = True
        Else
            Win = False
    End If
    
    Select Case I
        Case 1 To 4
            Won = Bet * 5
        Case 5 To 16
            Won = Bet * 2
    End Select
    
    If Win = True Then
        Money = Money + Won
        PicMoney.Cls
        PicWon.Cls
        PicMoney.Print FormatCurrency(Money)
        PicWon.Print FormatCurrency(Won)
    
        Else
            Money = Money - Bet
            Won = 0
            PicMoney.Cls
            PicWon.Cls
            PicMoney.Print FormatCurrency(Money)
            PicWon.Print FormatCurrency(Won)
    End If
    
    If Money = 0 Then
        MsgBox "You Are Out of Money", , "Broke"
        CmdCashOut.Enabled = False
        CmdChangeBet.Enabled = False
        CmdSpin.Enabled = False
    End If
    
    Close #1
    
    
End Sub
