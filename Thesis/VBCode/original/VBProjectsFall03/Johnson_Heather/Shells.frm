VERSION 5.00
Begin VB.Form frmShells1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Shells"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clear"
      Height          =   855
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txthowmany 
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton cmddiscount 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Now lets calculate the total with the dicount!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.PictureBox picbottomv 
      Height          =   1935
      Left            =   4440
      Picture         =   "Shells.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox picMHS 
      Height          =   1695
      Left            =   8400
      Picture         =   "Shells.frx":39F2
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox picvneck 
      Height          =   1935
      Left            =   10800
      Picture         =   "Shells.frx":90CA
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox picstripes 
      Height          =   1695
      Left            =   6360
      Picture         =   "Shells.frx":D1BB
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optbottomv 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optstripes 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optmhs 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton optvneck 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   11520
      MaskColor       =   &H0080C0FF&
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdmainmenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back to the Ordering Form!!!!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.PictureBox pictotalshells 
      Height          =   3015
      Left            =   5280
      ScaleHeight     =   2955
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   5640
      Width           =   4935
   End
   Begin VB.CommandButton cmdtoalcost 
      BackColor       =   &H000080FF&
      Caption         =   "TOTAL COST FOR THE SHELLS"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Label lblhowmany 
      BackColor       =   &H0080C0FF&
      Caption         =   "How Many Shells Would You Like?"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   27
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblstep5 
      BackColor       =   &H0080C0FF&
      Caption         =   "STEP 5: Back to the Main Menu"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      TabIndex        =   26
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label lblstep4 
      BackColor       =   &H0080C0FF&
      Caption         =   "STEP 4: CALCULATE THE DISCOUNT"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   25
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label lblstep3 
      BackColor       =   &H0080C0FF&
      Caption         =   "STEP 3: CALCULATE THE  TOTAL COST"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   24
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label lblstep2 
      BackColor       =   &H0080C0FF&
      Caption         =   "STEP 2: CHOOSE HOW MANY SHELLS YOU WANT"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label lblstep1 
      BackColor       =   &H0080C0FF&
      Caption         =   "STEP 1: CHOOSE A SHELL"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblwhichshell 
      BackColor       =   &H0080C0FF&
      Caption         =   "Which Cheerleading Shell Would You Like"
      BeginProperty Font 
         Name            =   "Mathematica7"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   20
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label lblvneck 
      BackColor       =   &H0080C0FF&
      Caption         =   "V NECK BRAID SHELL"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblMHS 
      BackColor       =   &H0080C0FF&
      Caption         =   "MHS SHELL"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblstrips 
      BackColor       =   &H0080C0FF&
      Caption         =   "STRIPES SHELL"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblbottomv 
      BackColor       =   &H0080C0FF&
      Caption         =   "BOTTOM V SHELL"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblcostbottom 
      BackColor       =   &H0080C0FF&
      Caption         =   "$50.98 per shell"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblcoststripes 
      BackColor       =   &H0080C0FF&
      Caption         =   "$48.25 per shell"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblcostmhs 
      BackColor       =   &H0080C0FF&
      Caption         =   "$56.76 per shell"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblcostvneck 
      BackColor       =   &H0080C0FF&
      Caption         =   "$52.75 per shell"
      BeginProperty Font 
         Name            =   "Architect"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "frmShells1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Shells (Shells1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'this form will ask you what shell you would like to purchase
                 'ask you how many shells you would like
                 'show you your total cost before the discount
                 'show you yout total cost after the discount
                 
Option Explicit
Dim Shells As String, S As Single, Price As Single
Dim Discount As Single, HowMany As Integer

Private Sub cmdclear_Click()
pictotalshells.Cls 'clears the shells output box
End Sub

Private Sub cmdmainmenu_Click()
frmOrder1.Show 'shows the order form
frmShells1.Hide 'hides the shells form
End Sub

Private Sub cmdquit_Click()
End 'ends the form
End Sub
Private Sub cmdtoalcost_Click()
If optbottomv = True Then
        pictotalshells.Print "Your Shell is the Bottom V Shell" 'if you push the optbottomv button this will print
    ElseIf optstripes = True Then
        pictotalshells.Print "Your Shell is the Stripes Shell" 'if you push the optstripes button this will print
    ElseIf optmhs = True Then
        pictotalshells.Print "Your Shell is the MHS Shell" 'if you push the optmhs button this is will print
    ElseIf optvneck = True Then
        pictotalshells.Print "Your Shell is the V Neck Braid Shell" 'if you push the optvneck button this will print
End If

If optbottomv = True Then
    S = 50.98 'if you push the optbottimv button this is how much each shell will cost
ElseIf optstripes = True Then
    S = 48.25 'if you push the optstripes button this is how much each shell will cost
ElseIf optmhs = True Then
    S = 56.76 'if you push the optmhs button this is how much each shell will cost
ElseIf optvneck = True Then
    S = 52.75 'if you push the optvneck button this is how much each shell will cost
End If

HowMany = txthowmany.Text 'this is how many shells you want
TotalCost = HowMany * S 'the total cost is how many shells you need multiplied by what button you pushed
    pictotalshells.Print "Your total cost for shells is "; FormatCurrency(TotalCost, 2) 'this will print the total for the shells before the discount
cmddiscount.Enabled = True 'you can now click on the discount button
End Sub

Private Sub cmddiscount_Click()
pictotalshells.Print
Select Case TotalCost
    Case Is < 50 'this is if the total cost is less then 50
        pictotalshells.Print "You save 4%" 'you will save 4% if the total cost is less then 50
            TotalCost = TotalCost - (TotalCost * 0.04) 'figures out the new total cost with the dicount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case 51 To 100 'this is if the total cost is between 50 and 100
        pictotalshells.Print "You save 5%" 'you will save 5% if the total cost is between 50 and 100
            TotalCost = TotalCost - (TotalCost * 0.05) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case 101 To 200 'this is if the total cost is between 101 and 200
        pictotalshells.Print "You save 6%" 'you will save 6% if the total cost is between 101 and 200
            TotalCost = TotalCost - (TotalCost * 0.06) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case 201 To 300 'this is if the total cost is between 201 and 300
        pictotalshells.Print "You save 7%" 'you will save 7% if the total cost is between 201 and 300
            TotalCost = TotalCost - (TotalCost * 0.07) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case 301 To 400 'this is if the total cost is between 301 and 400
        pictotalshells.Print "You save 8%" 'you will save 8% if the total cost is between 301 and 400
            TotalCost = TotalCost - (TotalCost * 0.08) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case 401 To 500 'this is if the total cost is between 401 and 500
        pictotalshells.Print "You save 9%" 'you will save 9% if the total cost is between 401 and 500
            TotalCost = TotalCost - (TotalCost * 0.09) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
    Case Is > 501 'this is if the total cost is greater then 501
        pictotalshells.Print "You save 10%" 'you will save 10% if the total cost is greater then 501
            TotalCost = TotalCost - (TotalCost * 0.1) 'figures out the new total cost with the discount
        pictotalshells.Print "Your total is now "; FormatCurrency(TotalCost, 2) 'prints out what the new total cost is
End Select
End Sub


