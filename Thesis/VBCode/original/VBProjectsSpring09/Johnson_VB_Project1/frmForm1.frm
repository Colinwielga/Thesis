VERSION 5.00
Begin VB.Form frmForm1 
   BackColor       =   &H0000FF00&
   Caption         =   "Black Jack"
   ClientHeight    =   11340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17040
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecrease 
      BackColor       =   &H000000C0&
      Caption         =   "Decrease Bet"
      Enabled         =   0   'False
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdIncrease 
      BackColor       =   &H000000C0&
      Caption         =   "Increase Bet"
      Enabled         =   0   'False
      Height          =   855
      Left            =   360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   7
      Left            =   14040
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   22
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   6
      Left            =   12120
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   21
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   5
      Left            =   10200
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   4
      Left            =   8280
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   19
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   3
      Left            =   6360
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   18
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   2
      Left            =   4440
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   17
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picDealerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   1
      Left            =   2520
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   10
      Left            =   6360
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   15
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   9
      Left            =   4440
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   14
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   8
      Left            =   2520
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   7
      Left            =   14040
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   6
      Left            =   12120
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   11
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   5
      Left            =   10200
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   4
      Left            =   8280
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   3
      Left            =   6360
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   2
      Left            =   4440
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox picPlayerCard 
      BackColor       =   &H000000C0&
      Height          =   2295
      Index           =   1
      Left            =   2520
      ScaleHeight     =   2235
      ScaleWidth      =   1635
      TabIndex        =   6
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdStay 
      BackColor       =   &H000000C0&
      Caption         =   "Stay"
      Enabled         =   0   'False
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeal 
      BackColor       =   &H000000C0&
      Caption         =   "Deal"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000C0&
      Caption         =   "Shuffle and Start The Game!"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton cmdHonorForm 
      BackColor       =   &H000000C0&
      Caption         =   "View regular Grads"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdHit 
      BackColor       =   &H000000C0&
      Caption         =   "Hit ME!"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblDealerSum 
      Height          =   615
      Left            =   8280
      TabIndex        =   32
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblDealerTotal 
      Caption         =   "Dealer is Showing:"
      Height          =   375
      Left            =   8280
      TabIndex        =   31
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblPlayerSum 
      Height          =   615
      Left            =   8400
      TabIndex        =   30
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label lblYourTotal 
      Caption         =   "Your Total"
      Height          =   375
      Left            =   8400
      TabIndex        =   29
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label lblCash 
      Height          =   495
      Left            =   3600
      TabIndex        =   28
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblCurrentCash 
      Caption         =   "Current Cash"
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblBet 
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblCurrentBet 
      Caption         =   "Current Bet"
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cards(4 To 55) As Boolean   'Declares a boolean array called "cards" to allow each card to only be dealt once
Dim Deck As Integer             'Declares an integer variable to count how many cards have been dealt so far so that the deck can be reshuffled eventually (resets after reshuffle)
Dim PlayerAces As Integer       'Counts how many many aces the player has
Dim DealerAces As Integer       'Counts how many aces the dealer has
Dim PlayerSum As Integer        'Keeps track of the total value of all the cards the player has each hand
Dim DealerSum As Integer        'Keeps track of the total value of all the cards the dealer has each hand
Dim PlayerCards As Integer      'Keeps track of the total number of all the cards the player has each hand (also works with where to place each card picture)
Dim DealerCards As Integer      'Keeps track of the total number of all the cards the dealer has each hand so that the dealer never takes more than seven cards (also works with where to place each card picture)
Dim DealerHidden As Integer     'keeps track of the first card the dealer gets every hand so that the player wont see it until after the player hits the "stay" command
Dim r As Integer                'a simple random integer 4-55 (52 unique possibilities, each representing a unique card) that will be used repeatedly
Dim Bet As Integer              'allows the player to make a bet before the dealer deals and allows the game to feel more real with actual wagers

Private Sub cmdStart_Click()    'Starts a new game, or reshuffles the deck and continues old game
Dim k As Integer        'Declares a simple integer for setting the card places with the valueless image of the back of the card
Bet = 1                 'resets the bet to $1 at the start of every new deck
lblBet.Caption = FormatCurrency(Bet)    'shows the initial bet each shuffle
lblCash.Caption = FormatCurrency(Cash)  'shows the total amount of cash the player has left (initially set to 100 as a public integer on the main module)
lblDealerSum.Caption = 0
lblPlayerSum.Caption = 0                'resets the amount shown by each player each shuffle

For k = 4 To 55
    Cards(k) = False                    'resets the availibity of each card
Next k

For k = 1 To 10
    picPlayerCard(k).Picture = LoadPicture(App.Path & "\" & CardPix(56))
Next k                                  'places an image of a face down card in each card space for the player

For k = 1 To 7
    picDealerCard(k).Picture = LoadPicture(App.Path & "\" & CardPix(56))
Next k                                  'places an image of a face down card in each card space for the dealer

Deck = 52                               'resets the total number of cards left in the deck

cmdStart.Enabled = False
cmdDeal.Enabled = True
cmdHit.Enabled = False
cmdStay.Enabled = False
cmdIncrease.Enabled = True
cmdDecrease.Enabled = True              'Makes it so that the player can increase/decrease their bet or deal a new hand after every shuffle, but can not "hit", "stay", or reshuffle until after a new hand is actually dealt

End Sub

Private Sub cmdDeal_Click() 'Deals a new hand to the player and the dealer
Dim ACard As Boolean        'A simple boolean term that prevents the cards from being distributed out of order off the deal (i.e. the player will get exactly 1 card before the dealer gets exactly one card etc.)
Dim b As Integer            'a simple integer that works with setting all the cards face down
Dim DealerShown As Integer  'An integer designed to show the value of the dealers one face up card

cmdDeal.Enabled = False     'prevents the player from clicking "deal", again, until after they click "stay"
cmdIncrease.Enabled = False 'prevents the player from increasing their bet again, until after this hand is over
cmdDecrease.Enabled = False 'prevents the player from decreasing their bet again, until after this hand is over
PlayerCards = 0             'resets the total number of cards the player has
DealerCards = 0             'resets the total number of cards the dealer has
PlayerSum = 0               'resets the total value of the cards the player has
DealerSum = 0               'resets the total value of the cards the dealer has
PlayerAces = 0              'resets the total number of aces the player has
DealerAces = 0              'resets the total number of aces the dealer has

Randomize                   'An independent command line that makes the Rnd() function actually totally random

For b = 1 To 10
picPlayerCard(b).Picture = LoadPicture(App.Path & "\" & CardPix(56))
Next b                                              'These past 3 lines set each of the players cards face down
For b = 1 To 7
picDealerCard(b).Picture = LoadPicture(App.Path & "\" & CardPix(56))
Next b                                              'These past 3 lines set each of the dealers cards face down

Cash = Cash - Bet                                   'After the player makes thei bet and clicks deal, this subtracts the players bet from their total cash
lblCash.Caption = FormatCurrency(Cash)              'Shows the player how much cash they still have

Do While (PlayerCards + DealerCards) < 4            'Prevents more than 4 cards from being dealt on the deal
    ACard = False
    Do While Not ACard                              'allows the player to only get 1 card until the dealer gets a card
    r = CInt((51) * Rnd()) + 4                      'produces a random integer from 4 to 55 (52 unique possibilites) to represent a random card the player might get
        If Cards(r) = False Then                    'if this random card has not been used yet this shuffle, then the player can recieve this card
            Cards(r) = True                         'prevents this random card from being used again until the next shuffle
            PlayerCards = PlayerCards + 1           'adds one more card to the total number of cards the player has
            ACard = True                            'makes it so that the player can not recieve another card until the dealer does
            Deck = Deck - 1                         'subtracts 1 card from the deck so that the deck can be reshuffled eventually
            picPlayerCard(PlayerCards).Picture = LoadPicture(App.Path & "\" & CardPix(r))  'shows the card the player just got
            If r > 40 Then
                PlayerSum = PlayerSum + 10          'if the card represented by "r" is a "ten" card or a "face" card, then only 10 more points will be added to the value of the players hand
            ElseIf r <= 7 And PlayerSum <= 10 Then
                PlayerAces = PlayerAces + 1
                PlayerSum = PlayerSum + 11          'if the card is an ace and the players sum is less than or equal to 10, then the player gets +11 to their player sum
            Else
                PlayerSum = PlayerSum + Fix(r / 4)  'for all other cases, the value of r divided by 4, rounded down to the next integer is the value of the card added to the players hand
            End If
        End If
    Loop
    
    ACard = False
    Do While Not ACard                              'allows the dealer to only get 1 card until the player gets another card
    r = CInt((51) * Rnd()) + 4                      'produces a random integer from 4 to 55 (52 unique possibilites) to represent a random card the dealer might get
        If Cards(r) = False Then                    'if this random card has not been used yet this shuffle, then the dealer can recieve this card
            Cards(r) = True                         'prevents this random card from being used again until the next shuffle
            DealerCards = DealerCards + 1           'adds one more card to the total number of cards the dealer has
            ACard = True                            'makes it so that the dealer can not recieve another card until the player does
            Deck = Deck - 1                         'subtracts 1 card from the deck so that the deck can be reshuffled eventually
            If DealerCards = 1 Then
                DealerHidden = r                    'Keeps track of the dealers first card for later purposes
            ElseIf DealerCards = 2 Then
                picDealerCard(DealerCards).Picture = LoadPicture(App.Path & "\" & CardPix(r)) 'Shows the dealers second card
            End If
            If r > 40 Then
                DealerSum = DealerSum + 10          'if the card represented by "r" is a "ten" card or a "face" card, then only 10 more points will be added to the value of the dealers hand
                DealerShown = 10                    'if the value of the second card is ten, then that will be the value the player sees
            ElseIf r <= 7 And DealerSum <= 10 Then
                DealerAces = DealerAces + 1
                DealerSum = DealerSum + 11          'if the card is an ace and the players sum is less than or equal to 10, then the dealer gets +11 to their dealer sum
            Else
                DealerSum = DealerSum + Fix(r / 4)  'for all other cases, the value of r divided by 4, rounded down to the next integer is the value of the card added to the players hand
            End If
        End If
        
        If r <= 7 Then
            DealerShown = 11                        'if the second card is an ace, then the player will see 11 as the value of the dealers hand
        Else
            DealerShown = Fix(r / 4)                'for all other cases, the value of r divided by 4, rounded down to the next integer is the value the player can see
        End If
        
    Loop
Loop

If PlayerSum = 21 Then
    MsgBox "You got 21, you can't hit anymore", , "Good Job!"
    cmdHit.Enabled = False                          'prevents the player from taking more cards after they reach 21
Else
    MsgBox "Your total is:" & PlayerSum & ". You can choose to hit or stay", , "Status"
    cmdHit.Enabled = True                           'allows the player to "hit" after they have recieved their first two cards
End If

cmdStay.Enabled = True                              'whether or not the player can "hit", they can always "stay"

lblPlayerSum.Caption = PlayerSum
lblDealerSum.Caption = DealerShown                  'Shows the values of the cards on the table for the respective players

End Sub

Private Sub cmdIncrease_Click()                     'generally increases the players bet

If Bet >= 5 And Cash >= 5 Then
    MsgBox "You can not bet more than $5 on one hand", , "No can do"
    Bet = 5
    lblBet.Caption = FormatCurrency(Bet)            'Prevents the player from ever increasing their bet beyond $5 per hand
ElseIf Cash <= Bet And Cash >= 1 Then
    MsgBox "You can not bet more than what you have", , "No can do"
    Bet = Cash
    lblBet.Caption = FormatCurrency(Bet)            'Limits the bet to never be greater than the players remaining cash (if their remaining cash happens to be less than $5)
Else
    Bet = Bet + 1
    lblBet.Caption = FormatCurrency(Bet)            'for all other cases, allows the bet to be increased up to $5
End If

End Sub

Private Sub cmdDecrease_Click()                     'generally decreases the players bet

If Bet <= 1 Then
    MsgBox "You can not bet less than $1 on one hand", , "No can do"
    Bet = 1
    lblBet.Caption = FormatCurrency(Bet)            'Prevents the player from ever decreasing their bet below $1 per hand
Else
    Bet = Bet - 1
    lblBet.Caption = FormatCurrency(Bet)            'for all other cases, allows the bet to be decreased down to $1
End If
    
End Sub

Private Sub cmdHit_Click()                          'gives the player a new card
Dim Found As Boolean                                'Found is used in a "do while not" boolean loop below that prevents repeat cards to be dealt
Found = False                                       'allows a new card to be found every time the player clicks "hit"

Randomize                                           'An independent command line that makes the Rnd() function actually totally random

Do While Not Found                                  'Keep looking for a card until one is found that hasnt been used yet this shuffle

r = CInt((51) * Rnd()) + 4                          'produces a random integer from 4 to 55 (52 unique possibilites) to represent a random card the player might get
    If Cards(r) = False Then                        'if the card has not been used yet, that player can recieve it
        Cards(r) = True                             'if the player recieves the card, the card can not be used again until the deck gets reshuffled
        Found = True                                'do not look for anymore cards until the player clicks "hit" or "deal" again
        PlayerCards = PlayerCards + 1               'add 1 to the number of cards in the players hand
        Deck = Deck - 1                             'subtract 1 from the total number of cards in the deck
        picPlayerCard(PlayerCards).Picture = LoadPicture(App.Path & "\" & CardPix(r))   'display the card in the next available spot for the player
        If r > 40 Then
            PlayerSum = PlayerSum + 10              'if the card represented by "r" is a "ten" card or a "face" card, then only 10 more points will be added to the value of the players hand
        ElseIf r <= 7 And PlayerSum <= 10 Then
            PlayerAces = PlayerAces + 1
            PlayerSum = PlayerSum + 11              'if the card is an ace and the players sum is less than or equal to 10, then the player gets +11 to their player sum
        Else
            PlayerSum = PlayerSum + Fix(r / 4)      'for all other cases, the value of r divided by 4, rounded down to the next integer is the value of the card added to the players hand
        End If
    End If
Loop

If PlayerSum > 21 And PlayerAces = 0 And Deck > 20 Then
    MsgBox "You lose this hand", , "Bust!"
    cmdHit.Enabled = False
    cmdStay.Enabled = False
    cmdDeal.Enabled = True
    cmdIncrease.Enabled = True
    cmdDecrease.Enabled = True                      'If the player gets more than 21 points without any aces, then they have to deal a new hand
ElseIf PlayerSum > 21 And PlayerAces = 0 And Deck < 20 Then
    MsgBox "You lose this hand", , "Bust!"
    cmdHit.Enabled = False
    cmdStay.Enabled = False
    cmdDeal.Enabled = False
    cmdIncrease.Enabled = False
    cmdDecrease.Enabled = False
    MsgBox "The deck needs to be reshuffled now. Please click the shuffle button", , "Reshuffle"
    cmdStart.Enabled = True                         'If the player gets more than 21 points without any aces and the deck has less than 20 cards, then the dealer has to reshuffle
ElseIf PlayerSum > 21 And PlayerAces > 0 Then
    PlayerSum = PlayerSum - 10
    PlayerAces = PlayerAces - 1                     'If the player gets more than 21, but has an ace that has a value of 11, then the ace gets reduced to 1
End If

If PlayerSum = 21 Then
    MsgBox "You got 21, you can't hit anymore", , "Good Job!"
    cmdHit.Enabled = False                          'if the player gets exactly 21, then the player has to "stay" and see what the dealer gets
ElseIf PlayerSum < 21 Then
    MsgBox "Your total is:" & PlayerSum & ". You can choose to hit or stay", , "Status" 'If the player has less than 21 after they hit, then they can "hit" again or "stay"
End If

lblPlayerSum.Caption = PlayerSum                    'shows the total value of the players hand

End Sub

Private Sub cmdStay_Click()                         'ends the players turn and causes the dealer to start evaluating their own hand
cmdHit.Enabled = False
cmdStay.Enabled = False                             'prevents the player from clicking "hit" or "stay" again until a new hand is dealt

picDealerCard(1).Picture = LoadPicture(App.Path & "\" & CardPix(DealerHidden))  'Shows the dealers hidden card for the player to see

Do While DealerCards < 6 And DealerSum < 17 And DealerSum < PlayerSum       'only allows the dealer to take a new card if it has less than 6 cards, less than 17 points, and less points than the player
r = CInt((51) * Rnd()) + 5                          'produces a random integer from 4 to 55 (52 unique possibilites) to represent a random card the dealer might get
    If Cards(r) = False Then                        'if the card has not been used yet, that dealer can recieve it
        Cards(r) = True                             'if the dealer recieves the card, the card can not be used again until the deck gets reshuffled
        DealerCards = DealerCards + 1               'add 1 to the number of cards in the dealers hand
        Deck = Deck - 1                             'subtract 1 from the total number of cards in the deck
        picDealerCard(DealerCards).Picture = LoadPicture(App.Path & "\" & CardPix(r))   'show the card that the dealer gets
        If r > 40 Then
            DealerSum = DealerSum + 10              'if the card represented by "r" is a "ten" card or a "face" card, then only 10 more points will be added to the value of the dealers hand
        ElseIf r <= 7 And DealerSum <= 10 Then
            DealerAces = DealerAces + 1
            DealerSum = DealerSum + 11              'if the card is an ace and the dealers sum is less than or equal to 10, then the dealer gets +11 to their player sum
        Else
            DealerSum = DealerSum + Fix(r / 4)      'for all other cases, the value of r divided by 4, rounded down to the next integer is the value of the card added to the players hand
        End If
    End If
    
    Do While DealerSum > 21 And DealerAces > 0
        DealerSum = DealerSum - 10                  'If the dealer ever has more than 21 points and an ace that is valued at 11 points, then subtract 10 points from the dealers total
        DealerAces = DealerAces - 1
    Loop
Loop

lblDealerSum.Caption = DealerSum                    'show the dealer's sum for the player to see

MsgBox "The Dealers total is: " & DealerSum & ". Your total is: " & PlayerSum, , "Totals"   'give the player an update of the score after the dealer is done

If DealerSum > 21 Then
    MsgBox "The dealer has busted. You win this hand! Now let it ride", , "Try again"
    Cash = Cash + (Bet * 2)                         'If the dealer busts, the player gets twice their bet back
    lblCash.Caption = FormatCurrency(Cash)          'show the player their new cash balance
ElseIf DealerSum >= PlayerSum And DealerSum <= 21 Then
    MsgBox "You lose this hand. Here's a free drink, on the house.", , "Try again"
ElseIf PlayerSum = 21 And DealerSum < PlayerSum Then
    MsgBox "You win this hand! Good Job, now let it  ride", , "Try again"
    Cash = Cash + (Bet * 5)                         'If the player gets 21 and wins, the player gets five times their bet back
    lblCash.Caption = FormatCurrency(Cash)          'show the player their new cash balance
Else
    MsgBox "You win this hand! Good Job, now let it  ride", , "Try again"
    Cash = Cash + (Bet * 3)                         'If the player gets 21 and wins, the player gets three times their bet back
    lblCash.Caption = FormatCurrency(Cash)          'show the player their new cash balance
End If

If Deck <= 20 Then
    MsgBox "The deck needs to be reshuffled now. Please click the shuffle button", , "Reshuffle"
    cmdStart.Enabled = True
    cmdDeal.Enabled = False
    cmdIncrease.Enabled = False
    cmdDecrease.Enabled = False                     'If the deck has less than 20 cards left, the dealer has to reshuffle before a new hand can be dealt
Else
    cmdDeal.Enabled = True
    cmdIncrease.Enabled = True
    cmdDecrease.Enabled = True                      'otherwise, once this hand is over, a new hand can be dealt
End If

End Sub

Private Sub cmdHonorForm_Click()

frmRegularGrads.Show                                'go back to the main form
frmForm1.Hide                                       'hide black jack

End Sub

Private Sub cmdQuit_Click()
    End                                             'end program
End Sub

