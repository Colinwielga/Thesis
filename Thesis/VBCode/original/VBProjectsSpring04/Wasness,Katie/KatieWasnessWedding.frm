VERSION 5.00
Begin VB.Form frmWedding 
   BackColor       =   &H8000000D&
   Caption         =   "BUDGET"
   ClientHeight    =   8580
   ClientLeft      =   4950
   ClientTop       =   3450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   9885
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results Box"
      Height          =   1095
      Left            =   3840
      TabIndex        =   12
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdTotalCost 
      Caption         =   "Click Here to Find Out the Total Cost of Your Wedding"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCostofReception 
      Caption         =   "Click To Display the Cost of the Reception"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdCostofCeremony 
      Caption         =   "Click To Display the Cost of the Ceremony"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdWPCost 
      Caption         =   "Click To Display Cost of the Wedding Party Attire"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdCostFlowers 
      Caption         =   "Click To Display Cost of Flowers"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   9315
      TabIndex        =   5
      Top             =   2640
      Width           =   9375
   End
   Begin VB.CommandButton cmdReception 
      Caption         =   "Choose Reception Site"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdCeremony 
      Caption         =   "Choose Ceremony Site"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdWeddingParty 
      Caption         =   "Choose Your Wedding Party and Their Attire"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdFlowers 
      BackColor       =   &H80000013&
      Caption         =   " Choose Flowers"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2640
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   7800
      TabIndex        =   0
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label labtitle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "YOUR WEDDING BUDGET"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "frmWedding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmWedding(Katie Wasness Wedding)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of the overall program is to assist an individual in putting together a basic budget for their wedding.
'Purpose--This form is designed to be the basic menu.  From here the individual can select different elements that are in a wedding to decide on prices.
'           Also, on this form one can choose to print the final budget.
'Purpose--The purpose of modOverall is to be able to transfer the numbers from one form to another, so that at the end the user may see the whole budget.
Dim Difference As Single

Private Sub cmdCeremony_Click()
'This button is used to open up the Ceremony form.
frmCeremony.Show
frmWedding.Hide
cmdCostofCeremony.Enabled = True
cmdReception.Enabled = True
End Sub


Private Sub cmdClear_Click()
'This button is to Clear the Results.
picResults.Cls
End Sub

Private Sub cmdCostFlowers_Click()
'This button is used to display the cost of the Flowers
picResults.Print ""
picResults.Print Bride; " and "; Groom; " have decided to have "; TotalFlowers; WhichFlower; " in their wedding."
picResults.Print "The total cost for those flowers is "; FormatCurrency(CostOfFlowers); "."
End Sub

Private Sub cmdCostofCeremony_Click()
'This button is used to display the cost of the Ceremony
picResults.Print ""
picResults.Print "The Cost of Your Ceremony in "; Choice; " is "; FormatCurrency(CostOfCeremony); "."
End Sub

Private Sub cmdCostofReception_Click()
'This button is used to display the cost of the Reception.
picResults.Print ""
picResults.Print Bride; " and "; Groom; " have decided to have the reception at "; Choice2
picResults.Print "          and the total cost of the reception will be "; FormatCurrency(CostOfReception); "."
End Sub

Private Sub cmdFlowers_Click()
'This button is used to see the Flowers form.
frmFlowers.Show
frmWedding.Hide
cmdCostFlowers.Enabled = True
cmdCeremony.Enabled = True
End Sub

Private Sub cmdQuit_Click()
'This button is to End the program.
End
End Sub

Private Sub cmdReception_Click()
'this button is to see the Reception Form.
frmReception.Show
frmWedding.Hide
cmdCostofReception.Enabled = True
cmdTotalCost.Enabled = True
End Sub

Private Sub cmdTotalCost_Click()
'This button is to display the total cost of the wedding and see if it is under the projected amount
picResults.Print ""
TotalCostofWedding = TotalCostofAttire + CostOfFlowers + CostOfCeremony + CostOfReception
picResults.Print "The Total Cost of Your Wedding is "; FormatCurrency(TotalCostofWedding); "."
If Budget > TotalCostofWedding Then
    picResults.Print "Good Job, You Stayed Within Your Projected Budget of "; FormatCurrency(Budget); "."
Else
    Difference = TotalCostofWedding - Budget
    picResults.Print "Uh Oh! "; Bride; " and "; Groom; " are over your projected budget by "; FormatCurrency(Difference); ".  You Better Find Some More Money."
End If
picResults.Print "Have a great and wonderful day.  Thank you for using this program."
End Sub

Private Sub cmdWeddingParty_Click()
'this button is to see the Wedding Party Form
FrmWeddingParty.Show
frmWedding.Hide
cmdWPCost.Enabled = True
cmdFlowers.Enabled = True
End Sub

Private Sub cmdWPCost_Click()
'this button is to display the cost of the Wedding Party Attire
If TotalCostofAttire = 0 Then
    TotalCostofAttire = CostofDress + TotalCostBM + CostofTux + TotalCostGMU
End If
picResults.Print ""
picResults.Print "The Total Cost of the Wedding Party Attire is "; FormatCurrency(TotalCostofAttire); "."
picResults.Print "The Cost of the Bride's Attire is "; FormatCurrency(CostofDress); "."
picResults.Print "The Cost of the Groom's Attire is "; FormatCurrency(CostofTux); "."
picResults.Print "The Total Cost of the Bridesmaid's Attire is "; FormatCurrency(TotalCostBM); "."
picResults.Print "The Total Cost of the Groomsmen and Usher's Attire is "; FormatCurrency(TotalCostGMU)

End Sub



Private Sub Form_Load()
' These are to congratulate the user and to get teh name of the Bride, Groom, and the projected budget.
MsgBox "Congratulations!! You are getting married.  This program is designed to help you plan your budget.", , "Congratulations!"
Bride = InputBox("What is the bride's name?", "Bride's Name")
Groom = InputBox("What is the groom's name?", "Groom's Name")
Budget = InputBox("What is the projected budget for your wedding in dollars and cents using a decimal point?", "Wedding Budget")

End Sub

