VERSION 5.00
Begin VB.Form frmMBrown 
   BackColor       =   &H00404040&
   Caption         =   "Brunette Male"
   ClientHeight    =   9255
   ClientLeft      =   1110
   ClientTop       =   870
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   13125
   Begin VB.CommandButton cmdSneakers 
      BackColor       =   &H00FF0000&
      Caption         =   "Sneakers - $65"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      BackColor       =   &H00404040&
      Caption         =   "Boots - $120"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   14.25
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
   Begin VB.CommandButton cmdSandles 
      BackColor       =   &H00004080&
      Caption         =   "Sandles - $25"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
   Begin VB.CommandButton cmdDressShirt 
      BackColor       =   &H0000FFFF&
      Caption         =   "Shirt and Tie -$80"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "T Shirt - $15"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
   Begin VB.CommandButton cmdPolo 
      BackColor       =   &H00FF0000&
      Caption         =   "Polo Shirt - $50"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
   Begin VB.CommandButton cmdSlack 
      BackColor       =   &H00808080&
      Caption         =   "Slacks - $95"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      BackColor       =   &H00008000&
      Caption         =   "Shorts - $60"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      Caption         =   "Jeans - $80"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
         Name            =   "Niagara Solid"
         Size            =   14.25
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
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      BackColor       =   &H00C0FFFF&
      Caption         =   "Display Your Name!"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      MaskColor       =   &H00C0FFFF&
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
      Left            =   8760
      ScaleHeight     =   675
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton cmdSatisfy 
      BackColor       =   &H0080FFFF&
      Caption         =   "I'm Satisfied With My Outfit"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdNothing1 
      Caption         =   "Buy Nothing"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   14.25
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
         Name            =   "Niagara Solid"
         Size            =   14.25
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
         Name            =   "Niagara Solid"
         Size            =   14.25
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
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   1560
      TabIndex        =   18
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label lblFunds 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "$$Your Funds$$"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   2280
      Width           =   3255
   End
End
Attribute VB_Name = "frmMBrown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Funds As Single
Dim Funds2 As Single
Dim Clothes(1 To 10) As String
Dim Prices(1 To 10) As Single
Dim CTR As Integer
Dim Found As Boolean
Dim PicName1 As String
Dim PicName2 As String



Private Sub cmdBack_Click()
'This Subroutine will move the user back a step, refunding the money from their last spending
'It will also take away the last letter of string put onto the PicName variable to avoid overload and error
'It makes the user only interact with the leg covering options again
cmdBack.Visible = False
cmdJeans.Visible = True
cmdSlack.Visible = True
cmdShorts.Visible = True
cmdDressShirt.Visible = False
cmdPolo.Visible = False
cmdTshirt.Visible = False
cmdNothing1.Visible = False
cmdNothing.Visible = True

PicName1 = PicName 'updates the PicName Variable for use of the back function

picFunds.Cls
picFunds.Print FormatCurrency(Funds)
End Sub

Private Sub cmdback2_Click()
'Same as for the Back subroutine except for the next step
cmdBack.Visible = True
cmdback2.Visible = False
cmdSandles.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdPolo.Visible = True
cmdNothing1.Visible = True
cmdNothing2.Visible = False
cmdDressShirt.Visible = True
cmdTshirt.Visible = True

PicName2 = PicName1

picFunds.Cls
picFunds.Print FormatCurrency(Funds2)
End Sub

Private Sub cmdBoots_Click()
'this subroutine finds the value of boots and subtracts it from the total funds
'It then prints the total funds for the user to see
'It hides the foot options and enables the Satisfy button to allow the user to move on to the next form
Dim pos As Integer

Found = False



Do While (pos <= CTR) And (Found = False) 'Finds Boot's Value
    pos = pos + 1
    If Clothes(pos) = "Boots" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos) 'Subtracts from total funds

If Funds3 < 0 Then 'Doesn't allow Overdrawing
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    cmdNothing2.Visible = False
    cmdSneakers.Visible = False
    cmdBoots.Visible = False
    cmdSandles.Visible = False
    cmdSatisfy.Visible = True
    PicName3 = PicName2 + "B"
End If

End Sub

Private Sub cmdDressShirt_Click()
'this subroutine finds the value of the dress shirt and subtracts it from the total funds
'It adds a "D" to the PicName variable for the dress shirt and tie
'It makes it so only the foot options are available
'It Updates the PicName Variable to allow for the back function
cmdBack.Visible = False
cmdback2.Visible = True
cmdPolo.Visible = False
cmdDressShirt.Visible = False
cmdTshirt.Visible = False
cmdSandles.Visible = True
cmdSneakers.Visible = True
cmdBoots.Visible = True
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Dim pos As Integer

PicName2 = PicName1 + "D"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "DressShirt" Then
        Found = True
    End If
Loop

picFunds.Cls
Funds3 = Funds2 - Prices(pos)
picFunds.Print FormatCurrency(Funds3)

End Sub

Private Sub cmdJeans_Click()
'this subroutine finds the value of the jeans and subtracts it from the total funds
'It adds a "J" to the PicName variable for Jeans
'Lastly It makes it so only the chest options are available
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSlack.Visible = False
cmdShorts.Visible = False
cmdPolo.Visible = True
cmdDressShirt.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

Dim pos As Integer

PicName1 = PicName + "J"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Jeans" Then
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
cmdJeans.Visible = True
cmdShorts.Visible = True
cmdSlack.Visible = True
cmdNothing.Visible = True
cmdName.Visible = False

Funds = 200
lblName = Ident

picFunds.Print FormatCurrency(Funds)

Open App.Path & "\GuyPrices.txt" For Input As #1

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
cmdSlack.Visible = False
cmdShorts.Visible = False
cmdPolo.Visible = False
cmdDressShirt.Visible = False
cmdTshirt.Visible = False
cmdNothing.Visible = False
cmdNothing1.Visible = True

PicName1 = PicName + "0" 'Adds a "0" for no clothing

Funds2 = Funds 'Updates Funds for Back function

End Sub

Private Sub cmdNothing1_Click()
'Similar to above, if the user selects this the only option they can prceed with is to again do nothing "no shirt, no shoes, no service"
cmdBack.Visible = False
cmdback2.Visible = True
cmdPolo.Visible = False
cmdDressShirt.Visible = False
cmdTshirt.Visible = False
cmdSandles.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdNothing1.Visible = False
cmdNothing2.Visible = True

PicName2 = PicName1 + "0"

Funds3 = Funds2
End Sub

Private Sub cmdNothing2_Click()
'Inputs a "0" to show no shoes purchased
'allows user to move to the congrats form
'updates PicName variable for the go back function
cmdNothing2.Visible = False
cmdSneakers.Visible = False
cmdBoots.Visible = False
cmdSandles.Visible = False
cmdSatisfy.Visible = True

PicName3 = PicName2 + "0"
End Sub

Private Sub cmdPolo_Click()
'Same as the Dress Shirt Routine except for a Polo
cmdBack.Visible = False
cmdback2.Visible = True
cmdPolo.Visible = False
cmdDressShirt.Visible = False
cmdTshirt.Visible = False
cmdSandles.Visible = True
cmdSneakers.Visible = True
cmdBoots.Visible = True
cmdNothing1.Visible = False
cmdNothing2.Visible = True

Dim pos As Integer

PicName2 = PicName1 + "P"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Polo" Then
        Found = True
    End If
Loop

picFunds.Cls
Funds3 = Funds2 - Prices(pos)
picFunds.Print FormatCurrency(Funds3)

End Sub

Private Sub cmdSandles_Click()
'Same as the Boots routine except for with Sandles
Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Sandles" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos)

If Funds3 < 0 Then 'doesnt let the user spend money they dont have
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    PicName3 = PicName2 + "S"
    cmdNothing2.Visible = False
    cmdSneakers.Visible = False
    cmdBoots.Visible = False
    cmdSandles.Visible = False
    cmdSatisfy.Visible = True
End If

End Sub

Private Sub cmdSatisfy_Click()
'When the the user has made their final selections they can move on to the next form using this subroutine
cmdName.Visible = True
frmMBlonde.Hide
frmCongrats.Show
End Sub

Private Sub cmdShorts_Click()
'Same as the Jeans routine except for with Shorts
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSlack.Visible = False
cmdShorts.Visible = False
cmdPolo.Visible = True
cmdDressShirt.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

Dim pos As Integer

PicName1 = PicName + "S"

Found = False

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

Private Sub cmdSlack_Click()
'Same as the Jeans routine except for with Slacks
cmdBack.Visible = True
cmdJeans.Visible = False
cmdSlack.Visible = False
cmdShorts.Visible = False
cmdPolo.Visible = True
cmdDressShirt.Visible = True
cmdTshirt.Visible = True
cmdNothing.Visible = False
cmdNothing1.Visible = True

Dim pos As Integer

PicName1 = PicName + "D"

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Slacks" Then
        Found = True
    End If
Loop

picFunds.Cls
Funds2 = Funds - Prices(pos)
picFunds.Print FormatCurrency(Funds2)
End Sub

Private Sub cmdSneakers_Click()
'Same as the boots routine except for with Sneakers
Dim pos As Integer

Found = False

Do While (pos <= CTR) And (Found = False)
    pos = pos + 1
    If Clothes(pos) = "Sneakers" Then
        Found = True
    End If
Loop

Funds3 = Funds3 - Prices(pos)

If Funds3 < 0 Then 'doesn't let the user spend money they don't have
    MsgBox ("You Don't have Enough Money for That"), , ("Error")
    Funds3 = Funds3 + Prices(pos)
Else
    picFunds.Cls
    picFunds.Print FormatCurrency(Funds3)
    cmdNothing2.Visible = False
    cmdSneakers.Visible = False
    cmdBoots.Visible = False
    cmdSandles.Visible = False
    cmdSatisfy.Visible = True
    PicName3 = PicName2 + "T"
End If
End Sub

Private Sub cmdTshirt_Click()
'Same as the Dress Shirt Routine except for with a Tshirt
cmdBack.Visible = False
cmdback2.Visible = True
cmdPolo.Visible = False
cmdDressShirt.Visible = False
cmdTshirt.Visible = False
cmdSandles.Visible = True
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
Funds3 = Funds2 - Prices(pos)
picFunds.Print FormatCurrency(Funds3)
End Sub




