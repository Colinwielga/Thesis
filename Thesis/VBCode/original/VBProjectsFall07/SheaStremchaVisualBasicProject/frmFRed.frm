VERSION 5.00
Begin VB.Form frmFRed 
   BackColor       =   &H00404040&
   Caption         =   "Red Headed Female"
   ClientHeight    =   9270
   ClientLeft      =   1320
   ClientTop       =   1095
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   13170
   Begin VB.CommandButton cmdSneakers 
      BackColor       =   &H000000FF&
      Caption         =   "Sneakers - $50"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdBoots 
      BackColor       =   &H00004080&
      Caption         =   "Hiking Boots - $120"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdHighHeels 
      BackColor       =   &H00808080&
      Caption         =   "High Heels - $20"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdJacket 
      BackColor       =   &H00C000C0&
      Caption         =   "Jacket - $80"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdTshirt 
      BackColor       =   &H00008000&
      Caption         =   "T Shirt - $15"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSweater 
      BackColor       =   &H00808080&
      Caption         =   "Sweater - $55"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSkirt 
      BackColor       =   &H000000FF&
      Caption         =   "Skirt - $50"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdShorts 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shorts - $40"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdJeans 
      BackColor       =   &H00FF8080&
      Caption         =   "Jeans - $70"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000080&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdback2 
      BackColor       =   &H00000080&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H0080FFFF&
      Caption         =   "Display Your Name!"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   3135
   End
   Begin VB.PictureBox picFunds 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   735
      Left            =   8520
      ScaleHeight     =   675
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdSatisfy 
      BackColor       =   &H0080FFFF&
      Caption         =   "I'm Satisfied With My Outfit"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdNothing1 
      Caption         =   "Buy Nothing"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNothing2 
      Caption         =   "Buy Nothing"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNothing 
      Caption         =   "Buy Nothing"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   1440
      TabIndex        =   18
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblFunds 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "$$Your Funds$$"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   2040
      Width           =   3255
   End
End
Attribute VB_Name = "frmFRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Funds As Single
Dim Funds2 As Single
Dim Clothes(1 To 10) As String
Dim PicName1 As String
Dim PicName2 As String
Dim Prices(1 To 10) As Single
Dim CTR As Integer
Dim Found As Boolean


Private Sub cmdBack_Click()
'This Subroutine will move the user back a step, refunding the money from their last spending
'It will also take away the last letter of string put onto the PicName variable to avoid overload and error
'It makes the user only interact with the leg covering options again
cmdBack.Visible = False
cmdJeans.Visible = True
cmdSkirt.Visible = True
cmdShorts.Visible = True
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdNothing.Visible = True
cmdNothing1.Visible = False

picFunds.Cls 'clears funds box to avoid many values listed

PicName1 = PicName 'reverts variable to avoid code overload

picFunds.Print FormatCurrency(Funds) 'Refunds money to last step
End Sub

Private Sub cmdback2_Click()
'This Subroutine will move the user back a step, refunding the money from their last spending
'It will also take away the last letter of string put onto the PicName variable to avoid overload and error
'It makes the user only interact with the Chest covering options again
cmdback2.Visible = False
cmdBack.Visible = True 'Allows User to go back another step
cmdHighHeels.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdSweater.Visible = True
cmdJacket.Visible = True
cmdTshirt.Visible = True
cmdNothing2.Visible = False

PicName2 = PicName1  'reverts variable to avoid code overload

picFunds.Cls 'clears funds box to avoid many values listed
picFunds.Print FormatCurrency(Funds2) 'refunds money

End Sub

Private Sub cmdBoots_Click()
'this subroutine finds the value of boots and subtracts it from the total funds
'It also updates the image to show the character wearing boots
'Adds a "B" to the PicName variable to correspond with the image name for Boots
Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False) 'Locates price from array
    pos = pos + 1
    If Clothes(pos) = "Boots" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos) 'subtracts price from funds

If Funds3 < 0 Then  'if the user trys to spend more than they have they recieve and error message and have to rethink their selections
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else 'If sufficient funds are present all appropriate actions take place
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    PicName3 = PicName2 + "B"
    cmdSatisfy.Visible = True 'Can now go to the next form
    cmdSneakers.Visible = False
    cmdHighHeels.Visible = False
    cmdBoots.Visible = False
    cmdNothing2.Visible = False
End If

End Sub

Private Sub cmdHighHeels_Click()
'This Subroutine is the same as for boots except for HighHeels
Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "HighHeels" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos)

If Funds3 < 0 Then
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    PicName3 = PicName2 + "H"
    cmdSatisfy.Visible = True
    cmdSneakers.Visible = False
    cmdHighHeels.Visible = False
    cmdBoots.Visible = False
    cmdNothing2.Visible = False
End If
End Sub

Private Sub cmdJacket_Click()
'this subroutine finds the value of the Jacket and subtracts it from the total funds
'It also updates the image to show the character wearing a Jacket
'It gets rid of the chest buttons and displays the foot covering options
cmdBack.Visible = False
cmdback2.Visible = True
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdHighHeels.Visible = True
cmdSneakers.Visible = True
cmdBoots.Visible = True
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Dim pos As Integer

PicName2 = PicName1 + "J" 'changes the picture variable to allow for stepping backward in the future

Found = False

Do While (pos <= CTR) And (Found = False) 'finds the price of the jacket from the array
    pos = pos + 1
    If Clothes(pos) = "Jacket" Then
        Found = True
    End If
Loop
picFunds.Cls 'Clears fund box to avoid multiple price listings
Funds3 = Funds2 - Prices(pos) 'subtracts price of jacket from total
picFunds.Print FormatCurrency(Funds3) 'displays updated fund levels

End Sub

Private Sub cmdJeans_Click()
'this subroutine finds the value of jeans and subtracts it from the total funds
'It also updates the image to show the character wearing jeans
'It takes away the leg options and enables the chest options
'Adds a "J" to the PicName variable to correspond with the image name for Jeans
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSkirt.Visible = False
cmdShorts.Visible = False
cmdSweater.Visible = True
cmdJacket.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

PicName1 = PicName + "J" 'Changes variable to accomidate the go back function

Dim pos As Integer
Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "jeans" Then
        Found = True
    End If
Loop
picFunds.Cls
Funds2 = Funds - Prices(pos)
picFunds.Print FormatCurrency(Funds2)
End Sub

Private Sub cmdName_Click()
'This subroutine get the name from the first form and displays it for the user to see
'It also loads the prices from the file so that they can be accessed by the buttons
'Lastly it displays the initial funds for clothing purchase $200
cmdNothing.Visible = True
cmdJeans.Visible = True
cmdShorts.Visible = True
cmdSkirt.Visible = True
Funds = 200
lblName = Ident
cmdName.Visible = False
picFunds.Print FormatCurrency(Funds)

Open App.Path & "\GirlPrices.txt" For Input As #1

Do Until EOF(1)
    CTR = CTR + 1
    Input #1, Clothes(CTR), Prices(CTR)
Loop

Close #1
End Sub

Private Sub cmdNothing_Click()
'This routine is similar to the other leg options except it doesn't put any clothing on the legs thus no money is deducted but the variable is still updated to accomidate the go back function
'If this is selected the chest options except for the "do nothing" are disabled because it wouldn't make sense to wear a shirt with no pants
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSkirt.Visible = False
cmdShorts.Visible = False
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdNothing1.Visible = True
cmdNothing.Visible = False

Funds2 = Funds 'funds are updated for "go back"

PicName1 = PicName + "0" 'A zero is added to describe no clothing
End Sub

Private Sub cmdNothing1_Click()
'Similar to above, if the user selects this the only option they can prceed with is to again do nothing "no shirt, no shoes, no service"
cmdBack.Visible = False
cmdback2.Visible = True
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdHighHeels.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Funds3 = Funds2

PicName2 = PicName1 + "0"
End Sub

Private Sub cmdNothing2_Click()
'Inputs a "0" to show no shoes purchased
'allows user to move to the congrats form
'updates PicName variable for the go back function
cmdSatisfy.Visible = True
cmdNothing2.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdHighHeels.Visible = False

PicName3 = PicName2 + "0"
End Sub

Private Sub cmdSatisfy_Click()
'When the the user has made their final selections they can move on to the next form using this subroutine
'Resets form for reuse
cmdName.Visible = True
frmFBlonde.Hide
frmCongrats.Show

End Sub

Private Sub cmdShorts_Click()
'This subroutine is almost identical to the jeans routine except with shorts
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSkirt.Visible = False
cmdShorts.Visible = False
cmdSweater.Visible = True
cmdJacket.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

Dim pos As Integer

PicName1 = PicName + "S"

Found = False
pos = 0
Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Shorts" Then
        Found = True
    End If
Loop
picFunds.Cls
Funds2 = Funds - Prices(pos)
picFunds.Print FormatCurrency(Funds2)
End Sub

Private Sub cmdSkirt_Click()
'This subroutine is almost identical to the jeans routine except with a skirt
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSkirt.Visible = False
cmdShorts.Visible = False
cmdSweater.Visible = True
cmdJacket.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

PicName1 = PicName + "K"

Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "skirt" Then
        Found = True
    End If

Loop

picFunds.Cls
Funds2 = Funds - Prices(pos)
picFunds.Print FormatCurrency(Funds2)

End Sub

Private Sub cmdSneakers_Click()
''This subroutine is almost identical to the boots routine except with sneakers
Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "sneakers" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos)

If Funds3 < 0 Then 'if the user trys to spend more than they have they recieve and error message and have to rethink their selections
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    PicName3 = PicName2 + "T"
    cmdSatisfy.Visible = True
    cmdSneakers.Visible = False
    cmdHighHeels.Visible = False
    cmdBoots.Visible = False
    cmdNothing2.Visible = False
End If
End Sub

Private Sub cmdSweater_Click()
'This subroutine is almost identical to the jacket routine except with a sweater
cmdNothing2.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdHighHeels.Visible = False
cmdBack.Visible = False
cmdback2.Visible = True
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdHighHeels.Visible = True
cmdSneakers.Visible = True
cmdBoots.Visible = True
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Dim pos As Integer

PicName2 = PicName1 + "S"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "sweater" Then
        Found = True
    End If
Loop
picFunds.Cls
Funds3 = Funds2 - Prices(pos) 'updates funds
picFunds.Print FormatCurrency(Funds3)
End Sub

Private Sub cmdTshirt_Click()
'This subroutine is almost identical to the jacket routine except with a tshirt
cmdBack.Visible = False
cmdback2.Visible = True
cmdSweater.Visible = False
cmdJacket.Visible = False
cmdTshirt.Visible = False
cmdHighHeels.Visible = True
cmdSneakers.Visible = True
cmdBoots.Visible = True
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Dim pos As Integer

PicName2 = PicName1 + "T"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Tshirt" Then
        Found = True
    End If
Loop

picFunds.Cls
Funds3 = Funds2 - Prices(pos) 'updates funds
picFunds.Print FormatCurrency(Funds3)

End Sub




