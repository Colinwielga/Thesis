VERSION 5.00
Begin VB.Form frmTaxInput 
   BackColor       =   &H80000013&
   Caption         =   "Tax Input"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTaxInput.frx":0000
      Left            =   5640
      List            =   "frmTaxInput.frx":000A
      TabIndex        =   39
      Top             =   8400
      Width           =   1455
   End
   Begin VB.PictureBox pic10 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   38
      Top             =   9240
      Width           =   855
   End
   Begin VB.PictureBox pic9 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   37
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton CmdViewTax 
      Caption         =   "View tax"
      Height          =   375
      Left            =   9360
      TabIndex        =   36
      Top             =   7320
      Width           =   1215
   End
   Begin VB.PictureBox pic8 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   35
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "Total"
      Height          =   375
      Left            =   9360
      TabIndex        =   34
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox pic7 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   33
      Top             =   6840
      Width           =   855
   End
   Begin VB.PictureBox pic6 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   9360
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox pic5 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   28
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmTaxInput.frx":0021
      Left            =   7680
      List            =   "frmTaxInput.frx":002E
      TabIndex        =   27
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtAdd 
      Height          =   375
      Left            =   7680
      TabIndex        =   26
      Text            =   "0"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdTotalIncome 
      Caption         =   "Adjusted GI"
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox picAdjusted 
      Height          =   375
      Left            =   7680
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   24
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdTax10 
      Caption         =   "Calculate #10"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdPayment 
      Caption         =   "Send/Recieve Payment"
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox txtAccountNum 
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtRoutingNum 
      Height          =   285
      Left            =   2880
      TabIndex        =   18
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc5 
      Caption         =   "Calculate #5"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txt8 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Text            =   "0"
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txt7 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Text            =   "0"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txt3 
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Text            =   "0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txt2 
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Text            =   "0"
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000013&
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   42
      Top             =   9960
      Width           =   1455
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7440
      X2              =   8760
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   720
      X2              =   10680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   720
      X2              =   10680
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   720
      X2              =   8760
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   720
      X2              =   10680
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7440
      X2              =   8760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   720
      X2              =   10680
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000013&
      Caption         =   "By Brent Mergen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000013&
      Caption         =   "Form 1040EZ"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   40
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label lblEnter9 
      Caption         =   "Enter Box 9 info from W-2"
      Height          =   375
      Left            =   8640
      TabIndex        =   32
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblEnter2 
      Caption         =   "Enter Box 2 info from  W-2"
      Height          =   375
      Left            =   8640
      TabIndex        =   31
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblOwe 
      Caption         =   "12.  If line 10 is larger than line 9, subtract line 9 from line 10.  You owe this amount.  "
      Height          =   375
      Left            =   840
      TabIndex        =   21
      Top             =   9240
      Width           =   4335
   End
   Begin VB.Label Label14 
      Caption         =   "11c. Type:"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label lblAccout 
      Caption         =   "11d. Account Number"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label lblRouting 
      Caption         =   "11b. Routing Number"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Label lblRefund 
      Caption         =   "11a.  This is your refund."
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label lblTax 
      Caption         =   "10.  Figure out your Tax."
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label lblAdd 
      Caption         =   "9.  Add Lines 7 and 8a.  These are your total payments."
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   6840
      Width           =   6615
   End
   Begin VB.Label lblEIC 
      Caption         =   "8.  Earned Income Credit (EIC)."
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   6240
      Width           =   6615
   End
   Begin VB.Label lblTaxWithheld 
      Caption         =   "7.  Federal incometax withheld from Box 2 of your For(s) W-2."
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label lblTaxable 
      Caption         =   "6.  Subtract line 5 from line 4. If line 5 is larger than line 4, enter "" -0- ""  This is your taxable income."
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   5160
      Width           =   6615
   End
   Begin VB.Label lblCalc2 
      Caption         =   $"frmTaxInput.frx":004E
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   4320
      Width           =   6615
   End
   Begin VB.Label lblAGI 
      Caption         =   "4.  Added lines 1, 2, and 3. This is your ADJUSTED GROSS INCOME"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   6615
   End
   Begin VB.Label lblUnemploy 
      Caption         =   "3.  Unemployment compensation and Alaska Permanent Fund dividends."
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3360
      Width           =   6615
   End
   Begin VB.Label lblTotal 
      Caption         =   "2.  Taxable interest. If the total is over $1,500, you cannot use Form 1040EZ."
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   6615
   End
   Begin VB.Label lbl1 
      Caption         =   "1.  Wages, Salaries, and tips. This should be shown in box 1 of your Form(s) W-2. "
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   6615
   End
End
Attribute VB_Name = "frmTaxInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E-Z Tazes (Brent's E-ZTax Form VBProject.vbp)
'Tax Inputs (frmTaxInput)
'Brent Timothy Mergen
'24 March 2006
'This is your main page for calculating your tax return value, have fun.  Oh yeah, you'll need your W-2 use this.

Option Explicit
Dim part6 As Single
Dim total As Single
Dim part7, part8 As Single
Dim adding As Single
Dim endingtax As Single

Private Sub cmdCalc5_Click()
    frmTaxInput.Hide
    frmTax5.Show
End Sub

Private Sub cmdPayment_Click()
    routingnumber = txtRoutingNum.Text
    accountnumber = txtAccountNum.Text
    If Combo1 = "Savings" Then
        selection = "savings"
    Else
        selection = "checking"
    End If
    frmTaxOutput.Show
    frmTaxInput.Hide
    
End Sub

Private Sub cmdsubmit_Click()
    pic5.Cls
    pic6.Cls
    picAdjusted.Cls
    total = txt1 + txt2 + txt3
    picAdjusted.Print total
    If Combo2 = "Dependent" Then
        pic5.Print overalltotal
        part6 = overalltotal
    End If
    If Combo2 = "Single" Then
        pic5.Print "8200"
        part6 = 8200
    End If
    If Combo2 = "Married" Then
        pic5.Print "16400"
        part6 = 16400
    End If
    If total > part6 Then
        pic6.Print total - part6
        answer6 = (total - part6)
    Else
        pic6.Print "0"
        answer6 = 0
    End If
End Sub

Private Sub cmdTax10_Click()
    frmTaxInput.Visible = False
    frmTax10.Visible = True
End Sub

Private Sub cmdTaxOutput_Click()
    frmTaxInput.Visible = False
    frmTaxOutput.Visible = True
End Sub

Private Sub cmdtotal_Click()
    pic7.Cls
    part7 = txt7.Text
    part8 = txt8.Text
    adding = part7 + part8
    pic7.Print adding
End Sub


Private Sub cmdTotalIncome_Click()
    txt1 = txtAdd.Text
    txt2 = txt2.Text
    txt3 = txt3.Text
    picAdjusted.Cls
    total = txt1 + txt2 + txt3
    picAdjusted.Print total
End Sub

Private Sub cmdViewTax_Click()
    pic7.Cls
    pic8.Cls
    pic9.Cls
    pic10.Cls
    pic8.Print Overalltax
    part7 = txt7.Text
    part8 = txt8.Text
    adding = part7 + part8
    pic7.Print adding
    If adding > Overalltax Then
        endingtax = adding - Overalltax
        pic9.Print endingtax
    End If
    If Overalltax > adding Then
        endingtax = Overalltax - adding
        pic10.Print endingtax
    End If
End Sub

