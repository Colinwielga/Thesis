VERSION 5.00
Begin VB.Form frmLettering1 
   BackColor       =   &H00FF8080&
   Caption         =   "Lettering"
   ClientHeight    =   9090
   ClientLeft      =   3300
   ClientTop       =   3150
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   12150
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FF0000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmddiscount 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Now Lets Calculate the Total with the Discount"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   5400
      Picture         =   "Lettering.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1755
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   9600
      Picture         =   "Lettering.frx":15D0
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   3360
      Picture         =   "Lettering.frx":2D58
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   7440
      Picture         =   "Lettering.frx":43DA
      ScaleHeight     =   1275
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton optcrazy 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optbowtie 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optdiamond 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optbrody 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txthowmanyshells 
      Height          =   615
      Left            =   7680
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txthowmanyletters 
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdtotalletters 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TOTAL COST FOR THE LETTERS"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.PictureBox picresultsletters 
      Height          =   3255
      Left            =   5760
      ScaleHeight     =   3195
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   5160
      Width           =   3855
   End
   Begin VB.CommandButton cmdskirts 
      BackColor       =   &H00FF0000&
      Caption         =   "Lets go back to the Main Menu!!!"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Label lblstep6 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 6: BACK TO THE MAIN MENU"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   30
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label lblstep5 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 5: COMPUTE THE TOTAL COST WITH THE DISCOUNT"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label lblstep4 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 4: COMPUTE THE TOTAL COST FOR THE LETTERS"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   28
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label lblstep3 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 3: HOW MANY LETTERS WILL BE ON EACH SHELL?"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label lblstep2 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 2: HOW MANY SHELLS DID YOU PREVIOUSLY ORDER?"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label lblstep1 
      BackColor       =   &H00FF8080&
      Caption         =   "STEP 1: CHOOSE A STYLE OF FONT"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   25
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label lbllettering 
      BackColor       =   &H00FF8080&
      Caption         =   "What type of lettering would you like on your shells?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   23
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblcrazy 
      BackColor       =   &H00FF8080&
      Caption         =   "CRAZY"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblbowtie 
      BackColor       =   &H00FF8080&
      Caption         =   "BOWTIE"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lbldiamond 
      BackColor       =   &H00FF8080&
      Caption         =   "DIAMOND"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblbrody 
      BackColor       =   &H00FF8080&
      Caption         =   "BRODY"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblhowmanyshells 
      BackColor       =   &H00FF8080&
      Caption         =   "How many shells do you need letters on?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   18
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblhowmanyletters 
      BackColor       =   &H00FF8080&
      Caption         =   "How many letters on each shell?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label lblcostcrazy 
      BackColor       =   &H00FF8080&
      Caption         =   "$6.45 per letter"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblcostbowtie 
      BackColor       =   &H00FF8080&
      Caption         =   "$4.92 per letter"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblcostdiamond 
      BackColor       =   &H00FF8080&
      Caption         =   "$5.56 per letter"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblcostbrody 
      BackColor       =   &H00FF8080&
      Caption         =   "$5.11 per letter"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   13
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frmLettering1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Lettering (Lettering1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'this form will let you choose the lettering you want
                 'ask you how many shells you will need lettering on
                 'ask you how many letters are on each shell
                 'figure out your total cost before the discount
                 'figure out your actual total cost after the discount

Option Explicit
Dim Shells As Integer, Letter As String
Dim HowManyShells As Integer, HowManyLetters As Integer
Dim P As Single, Total As Single
Private Sub cmdclear_Click()
picresultsletters.Cls 'clears the output box for letters
End Sub

Private Sub cmdquit_Click()
End 'ends the program
End Sub

Private Sub cmdskirts_Click()
frmOrder1.Show 'shows the orderform
frmShells1.Hide 'hides the shells form
frmSkirts1.Hide 'hides the skirts form
frmLettering1.Hide 'hides the lettering form
End Sub

Private Sub cmdtotalletters_Click()
HowManyShells = txthowmanyshells.Text
HowManyLetters = txthowmanyletters.Text
If optcrazy = True Then
        picresultsletters.Print "Your font is Crazy" 'when you push the optcrazy button this tells you what font you picked
    ElseIf optbowtie = True Then
        picresultsletters.Print "Your font is Bowtie" 'when you push the optbowtie button this tells you what font you picked
    ElseIf optdiamond = True Then
        picresultsletters.Print "Your font is Diamond" 'when you push the optdiamond button this tells you what font you picked
    ElseIf optbrody = True Then
        picresultsletters.Print "Your font is Brody Script" 'when you puch the optbrody button this tells you what font you picked
End If
If optcrazy = True Then
    P = 6.45 'if you push the optcrazy button this is how much each letter will cost
ElseIf optbowtie = True Then
    P = 4.92 'if you push the optbowtie button this is how much each letter will cost
ElseIf optdiamond = True Then
    P = 5.56 'if you push the optdiamond button this is how much each letter will cost
ElseIf optbrody = True Then
    P = 5.11 'if you push the optbrody button this is how much each letter will cost
End If
Total = P * HowManyLetters * HowManyShells
    'multiplies the amout of letters you have times how many shells times how much each letter is
    picresultsletters.Print "Your cost before the discount is "; FormatCurrency(Total, 2) 'prints out what your total cost is before the discount
picresultsletters.Print 'prints out a plain line
cmddiscount.Enabled = True 'you can now push the discount button
End Sub

Private Sub cmddiscount_Click()
Select Case Total
    Case Is < 30 'if your total is less then $30
        picresultsletters.Print "You save 2%" 'you save 2%
            TotalCostLettering = Total - (Total * 0.02) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 31 To 40 'if your total is between 30 and 41
        picresultsletters.Print "You save 3%" 'you save 3%
            TotalCostLettering = Total - (Total * 0.03) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 41 To 50 'if your total is between 40 and 50
        picresultsletters.Print "You save 4%" 'you save 4%
            TotalCostLettering = Total - (Total * 0.04) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 51 To 60 'if your total is between 50 and 60
        picresultsletters.Print "You save 5%" 'you save 5%
            TotalCostLettering = Total - (Total * 0.05) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new toal
    Case 61 To 70 'if yout total is between 60 and 70
        picresultsletters.Print "You save 6%" 'you save 6%
            TotalCostLettering = Total - (Total * 0.06) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 71 To 80 'if yout total is between 70 and 80
        picresultsletters.Print "You save 7%" 'you save 7%
            TotalCostLettering = Total - (Total * 0.07) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 80 To 90 'if your total is between 80 and 90
        picresultsletters.Print "You save 8%" 'you save 8%
            TotalCostLettering = Total - (Total * 0.08) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case 91 To 100 'if your total is between 90 and 100
        picresultsletters.Print "You save 9%" 'you save 9%
            TotalCostLettering = Total - (Total * 0.09) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
    Case Is > 100 'if your total is greater then 100
        picresultsletters.Print "You save 10%" 'you save 10%
            TotalCostLettering = Total - (Total * 0.1) 'figures out your new total
                picresultsletters.Print "Your total cost is now "; FormatCurrency(TotalCostLettering, 2) 'prints your new total
End Select
End Sub
