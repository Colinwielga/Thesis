VERSION 5.00
Begin VB.Form FrmCosts 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdJump 
      Caption         =   "Click to figure out net income"
      Height          =   735
      Left            =   6600
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdVariable 
      Caption         =   "Click to calculate unit product cost under variable costing"
      Height          =   735
      Left            =   4320
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdAbsorption 
      Caption         =   "Click to calculate unit product cost under absorption costing"
      Height          =   735
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtFO 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtVO 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtDL 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtDM 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2040
      ScaleHeight     =   1995
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   6600
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblFMO 
      BackColor       =   &H80000009&
      Caption         =   "Input per unit Fixed Overhead Cost in dollars"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblVO 
      BackColor       =   &H80000009&
      Caption         =   "Input per unit Variable Overhead cost in dollars"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblDL 
      BackColor       =   &H80000009&
      Caption         =   "Input per unit cost of Direct Labor in dollars"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblDM 
      BackColor       =   &H80000009&
      Caption         =   "Input per unit cost of Direct Materials in dollars"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmCosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable vs. Absorption Cost Accounting(VBProject.vbp)
'FrmCosts(VBProject.frm)
'Mike Rakes, 3/10
'The purpose of this form is to get the per unit Direct Materials cost, direct labor cost,
'variable overhead per unit cost, and per unit fixed cost from the user and then from that data
'calculate the per unit cost of an item under absorption costing and also under variable costing.
'There is a brief explanation of why the per unit costs under each method differ from each other

Option Explicit
'This form will calculate the variable and absorption costs per unit
'Those numbers can be used to calculate the income under the two costing methods
Public VC As Single
Public AC As Single
Dim DM As Single, DL As Single, VO As Single
Public FO As Single

Private Sub cmdAbsorption_Click()
'This sets each variable equal to a textbox
DM = txtDM.Text
DL = txtDL.Text
VO = txtVO.Text
FO = txtFO.Text
'Adds up the 4 text boxes
AC = DM + DL + VO + FO
'Prints the data from AC
    picResults.Print "Under absorption costing the per unit cost of Product A is "; FormatCurrency(AC)
'Makes the GUI more user friendly by eliminating buttons as they are used
cmdAbsorption.Enabled = False
cmdVariable.Enabled = True

End Sub

Private Sub cmdJump_Click()
'Hides one form and jumps to the Income Form
FrmCosts.Hide
FrmIncome.Show
End Sub

Private Sub cmdQuit_Click()
'Quits out of program
End
End Sub

Private Sub cmdVariable_Click()
'Calculates Variable Cost per unit
VC = DM + DL + VO
'Prints the Data from VC and an explanation of why it differs from AC
    picResults.Print
    picResults.Print "Under variable costing, the fixed overhead per unit is not included in the total cost."
    picResults.Print "Therefore, the variable cost per unit of Product A is "; FormatCurrency(VC)
cmdVariable.Enabled = False
cmdJump.Enabled = True
End Sub

Private Sub Form_Load()
'Again, makes the GUI more user friendly by emilinating buttons that don't need to be
'pressed a second time
cmdJump.Enabled = False
cmdVariable.Enabled = False
End Sub
