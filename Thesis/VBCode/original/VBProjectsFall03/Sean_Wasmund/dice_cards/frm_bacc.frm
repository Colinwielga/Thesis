VERSION 5.00
Begin VB.Form frm_bacc 
   BackColor       =   &H00008000&
   Caption         =   "Baccarat by Sean Wasmund"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_start 
      Caption         =   "Start"
      Height          =   1215
      Left            =   3600
      TabIndex        =   13
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmd_clear_bets 
      Caption         =   "Clear Bets"
      Height          =   615
      Left            =   7920
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmd_stand 
      Caption         =   "Stand"
      Height          =   1455
      Left            =   8640
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmd_draw 
      Caption         =   "Draw"
      Height          =   1455
      Left            =   7440
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox pic_money 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmd_bet_25 
      Caption         =   "Bet $25"
      Height          =   615
      Left            =   8520
      TabIndex        =   7
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmd_bet_10 
      Caption         =   "Bet $10"
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmd_bet_5 
      Caption         =   "Bet $5"
      Height          =   615
      Left            =   8520
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmd_bet_1 
      Caption         =   "Bet $1"
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox pic_bet 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmd_deal 
      Caption         =   "Deal"
      Height          =   1335
      Left            =   7560
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmd_quit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmd_rules 
      Caption         =   "Rules"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lbl_money 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Your Money"
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   6480
      Width           =   855
   End
   Begin VB.Image img_player_1 
      Height          =   2175
      Left            =   2280
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image img_player_2 
      Height          =   2175
      Left            =   3960
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image img_dealer_3 
      Height          =   2175
      Left            =   5640
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image img_player_3 
      Height          =   2175
      Left            =   5640
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Image img_dealer_2 
      Height          =   2175
      Left            =   3960
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image img_dealer_1 
      Height          =   2175
      Left            =   2280
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frm_bacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cards_dealt(1 To 6) As Integer      'the possible six cards dealt from the deck
Dim cards(1 To 52) As Integer           'the deck
Dim card_pic(1 To 52) As String         'jpgs of the deck
Dim pip_value(1 To 52) As Integer       'value of the cards
Dim players_cards(1 To 3) As String     'cards delt to player
Dim players_sum As Integer              'sum of card's pip values
Dim dealers_cards(1 To 3) As String     'cards dealt to banker
Dim dealers_sum As Integer              'sum of card's pip values

Dim bet As Integer                      'amount bet
Dim PATH As String
Dim money As Integer                    'money the player has

Private Sub cmd_bet_1_Click()           'command to bet $1
cmd_deal.Enabled = True
bet = bet + 1

If bet > money Then                     'checks for sufficent funds
    MsgBox ("You don't have enough money.")
    bet = bet - 1
End If

pic_bet.Cls
pic_bet.Print FormatCurrency(bet)
End Sub

Private Sub cmd_bet_10_Click()          'command to bet $10
cmd_deal.Enabled = True
bet = bet + 10

If bet > money Then                     'checks for sufficent funds
    MsgBox ("You don't have enough money.")
    bet = bet - 10
End If

pic_bet.Cls
pic_bet.Print FormatCurrency(bet)
End Sub

Private Sub cmd_bet_25_Click()          'command to bet $25
cmd_deal.Enabled = True
bet = bet + 25

If bet > money Then                     'checks for sufficent funds
    MsgBox ("You don't have enough money.")
    bet = bet - 25
End If

pic_bet.Cls
pic_bet.Print FormatCurrency(bet)
End Sub

Private Sub cmd_bet_5_Click()           'command to bet $5
cmd_deal.Enabled = True
bet = bet + 5

If bet > money Then                     'checks for sufficent funds
    MsgBox ("You don't have enough money.")
    bet = bet - 5
End If

pic_bet.Cls
pic_bet.Print FormatCurrency(bet)
End Sub

Private Sub cmd_clear_bets_Click()      'clears and resets bet
cmd_deal.Enabled = False
bet = 0
pic_bet.Cls
End Sub

Private Sub cmd_deal_Click()            'starts round
Dim i As Integer
Dim j As Integer

cmd_deal.Enabled = False                'opens and locks out appropriate commands
cmd_draw.Enabled = True
cmd_stand.Enabled = True
cmd_bet_1.Enabled = False
cmd_bet_5.Enabled = False
cmd_bet_10.Enabled = False
cmd_bet_25.Enabled = False
cmd_clear_bets.Enabled = False

img_dealer_1.Picture = LoadPicture(none)    'clears image boxes
img_dealer_2.Picture = LoadPicture(none)
img_dealer_3.Picture = LoadPicture(none)
img_player_1.Picture = LoadPicture(none)
img_player_2.Picture = LoadPicture(none)
img_player_3.Picture = LoadPicture(none)

Open PATH & "thedeck.txt" For Input As #1   'opens file to be put into array

For i = 1 To 52
    Input #1, cards(i), pip_value(i), card_pic(i)   'put values into array
Next i

Close #1

img_dealer_1.Picture = LoadPicture(PATH & "card_back.jpg")  'shows bankers cards face down
img_dealer_2.Picture = LoadPicture(PATH & "card_back.jpg")

For i = 1 To 6
    cards_dealt(i) = cards(Int(52 * Rnd + 1))   'generates better random numbers
Next i



i = 1                                           'deals cards 1st to player, 2nd to banker
For j = 1 To 3
    players_cards(j) = cards_dealt(i)
    i = i + 1
    dealers_cards(j) = cards_dealt(i)
    i = i + 1
Next j

players_sum = 0
dealers_sum = 0

img_player_1.Picture = LoadPicture(PATH & card_pic(players_cards(1)))   'shows players cards
img_player_2.Picture = LoadPicture(PATH & card_pic(players_cards(2)))

players_sum = pip_value(players_cards(1)) + pip_value(players_cards(2)) 'adds pip values
dealers_sum = pip_value(dealers_cards(1)) + pip_value(dealers_cards(2))

If players_sum >= 10 Then                           'baccarat drops the 1st digit of a 2 digit number
    players_sum = players_sum - 10
End If

If dealers_sum >= 10 Then
    dealers_sum = dealers_sum - 10
End If

If players_sum = 9 And dealers_sum < 9 Then         'checks for natural 9 and awards win if banker does not match natural 9 on the deal
    MsgBox ("Natural 9:  You Win!")
    money = money + bet                             'bet is won if conditional is true
    img_dealer_1.Picture = LoadPicture(PATH & card_pic(dealers_cards(1)))
    img_dealer_2.Picture = LoadPicture(PATH & card_pic(dealers_cards(2)))
    cmd_draw.Enabled = False
    cmd_stand.Enabled = False
    cmd_bet_1.Enabled = True
    cmd_bet_5.Enabled = True
    cmd_bet_10.Enabled = True
    cmd_bet_25.Enabled = True
    cmd_clear_bets.Enabled = True
    bet = 0
    pic_bet.Cls
End If

If players_sum = 9 And dealers_sum = 9 Then         'if banker matches natural 9, tie
    MsgBox ("Natural 9:  Tie, all bets off.")
    img_dealer_1.Picture = LoadPicture(PATH & card_pic(dealers_cards(1)))
    img_dealer_2.Picture = LoadPicture(PATH & card_pic(dealers_cards(2)))
    cmd_draw.Enabled = False
    cmd_stand.Enabled = False
    cmd_bet_1.Enabled = True
    cmd_bet_5.Enabled = True
    cmd_bet_10.Enabled = True
    cmd_bet_25.Enabled = True
    cmd_clear_bets.Enabled = True
    bet = 0
    pic_bet.Cls
End If

If players_sum = 8 Then             'alerts player to their natural 8
    MsgBox ("Natural 8")
End If

pic_money.Cls                       'in the case of a natural 9, money is updated
pic_money.Print FormatCurrency(money)

End Sub

Private Sub cmd_draw_Click()        'command if player draws another card
cmd_draw.Enabled = False
cmd_stand.Enabled = False
cmd_bet_1.Enabled = True
cmd_bet_5.Enabled = True
cmd_bet_10.Enabled = True
cmd_bet_25.Enabled = True
cmd_clear_bets.Enabled = True

img_player_3.Picture = LoadPicture(PATH & card_pic(players_cards(3)))   'shows players final card
img_dealer_1.Picture = LoadPicture(PATH & card_pic(dealers_cards(1)))   'shows bankers cards
img_dealer_2.Picture = LoadPicture(PATH & card_pic(dealers_cards(2)))

players_sum = players_sum + pip_value(players_cards(3))                 'adds final card to players pip value sum

If players_sum >= 10 Then                        'drops 1st digit of a 2 digit number
    players_sum = players_sum - 10
End If

If players_sum > dealers_sum Then               'makes dealers decision to draw or win
    img_dealer_3.Picture = LoadPicture(PATH & card_pic(dealers_cards(3)))
    dealers_sum = dealers_sum + pip_value(dealers_cards(3))
End If

If dealers_sum >= 10 Then                       'drops 1st digit of 2 digit number for banker
    dealers_sum = dealers_sum - 10
End If

If players_sum > dealers_sum Then               'evaluates who wins
    MsgBox ("You Win!")
    money = money + bet
ElseIf players_sum = dealers_sum Then
    MsgBox ("Tie, all bets are off.")
ElseIf players_sum < dealers_sum Then
    MsgBox ("You Lose")
    money = money - bet
End If

bet = 0                                         'resets bet and updates money
pic_bet.Cls
pic_money.Cls
pic_money.Print FormatCurrency(money)
End Sub

Private Sub cmd_quit_Click()                    'quits baccarat back to dice roller
frm_roller.Show
frm_bacc.Hide
cmd_start.Visible = True

img_dealer_1.Picture = LoadPicture(none)        'clears image boxes
img_dealer_2.Picture = LoadPicture(none)
img_dealer_2.Picture = LoadPicture(none)
img_player_1.Picture = LoadPicture(none)
img_player_2.Picture = LoadPicture(none)
img_player_3.Picture = LoadPicture(none)
End Sub

Private Sub cmd_rules_Click()                   'shows rules form
frm_rules.Show
frm_bacc.Enabled = False
End Sub

Private Sub cmd_stand_Click()                   'command if player stands
cmd_deal.Enabled = False
cmd_draw.Enabled = False
cmd_stand.Enabled = False
cmd_bet_1.Enabled = True
cmd_bet_5.Enabled = True
cmd_bet_10.Enabled = True
cmd_bet_25.Enabled = True
cmd_clear_bets.Enabled = True

img_dealer_1.Picture = LoadPicture(PATH & card_pic(dealers_cards(1)))
img_dealer_2.Picture = LoadPicture(PATH & card_pic(dealers_cards(2)))

If players_sum > dealers_sum Then               'makes bankers decision to draw or win
    img_dealer_3.Picture = LoadPicture(PATH & card_pic(dealers_cards(3)))
    dealers_sum = dealers_sum + pip_value(dealers_cards(3))
End If

If dealers_sum >= 10 Then
    dealers_sum = dealers_sum - 10
End If

If players_sum > dealers_sum Then               'evaluates win
    MsgBox ("You Win!")
    money = money + bet
ElseIf players_sum = dealers_sum Then
    MsgBox ("Tie, all bets are off.")
ElseIf players_sum < dealers_sum Then
    MsgBox ("You Lose")
    money = money - bet
End If

bet = 0
pic_bet.Cls
pic_money.Cls
pic_money.Print FormatCurrency(money)
End Sub

Private Sub cmd_start_Click()               'initiates the game
money = 50                                  'sets starting money
pic_money.Cls
pic_money.Print FormatCurrency(money)

cmd_start.Visible = False
cmd_deal.Enabled = False
cmd_draw.Enabled = False
cmd_stand.Enabled = False
cmd_bet_1.Enabled = True
cmd_bet_5.Enabled = True
cmd_bet_10.Enabled = True
cmd_bet_25.Enabled = True
End Sub



Private Sub Form_Load()
PATH = "M:\cs130\sean_wasmund\dice_cards\cards\"        'path
Randomize                               'randomize random number clock

cmd_deal.Enabled = False
cmd_draw.Enabled = False
cmd_stand.Enabled = False
cmd_bet_1.Enabled = False
cmd_bet_5.Enabled = False
cmd_bet_10.Enabled = False
cmd_bet_25.Enabled = False

End Sub
