VERSION 5.00
Begin VB.Form frmWeddingParty 
   BackColor       =   &H8000000D&
   Caption         =   "WEDDING PARTY"
   ClientHeight    =   7950
   ClientLeft      =   4770
   ClientTop       =   3630
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10320
   Begin VB.CommandButton cmdshowAttireCost 
      Caption         =   "Show Attire Cost Total"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdAlphaGMU 
      Caption         =   "Click Here to List the Men in the Wedding Party with Groomsmen and Ushers in Alphabetical Order by Last Name"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   9
      Top             =   6120
      Width           =   3615
   End
   Begin VB.CommandButton cmdAlphaBM 
      Caption         =   "Click Here to List the Women in Wedding Party with the Bridesmaid's in Alphabetical Order by Last Name"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   6120
      Width           =   3615
   End
   Begin VB.PictureBox picMen 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   5280
      ScaleHeight     =   3795
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
   End
   Begin VB.PictureBox picWomen 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   600
      ScaleHeight     =   3795
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton cmdMen 
      Caption         =   "Click to Record the Names of your Groomsmen, Ushers, and Ring Bearer"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdWomen 
      Caption         =   "Click to Record the Names of your Bridesmaids, Flower Girl, and Personal Attendent"
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.PictureBox picAttire 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   7035
      TabIndex        =   3
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdAttire 
      Caption         =   "Choose Attire for Wedding Party"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackTo 
      Caption         =   "Go Back to Main Menu"
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   8880
      TabIndex        =   0
      Top             =   6840
      Width           =   855
   End
End
Attribute VB_Name = "FrmWeddingParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pjtWeddingBudget(Katie Wasness Wedding)
'frmWeddingParty(Katie Wasness Wedding Wedding Party)
'Author = Katie Wasness
'Date Written: March 10-15
'Purpose--The purpose of this form is to be the menu for choosing the who the bridal pary is and what they should wear.
Dim BridesmaidFirst(1 To 100) As String
Dim BridesmaidLast(1 To 100) As String
Dim MaidOfHonor As String
Dim FlowerGirl As String
Dim PersonalAttendent As String
Dim GroomsmenFirst(1 To 100) As String
Dim GroomsmenLast(1 To 100) As String
Dim BestMan As String
Dim RingBearer As String
Dim UsherFirst(1 To 100) As String
Dim UsherLast(1 To 100) As String



Private Sub cmdAlphaBM_Click()
'this button is used to alphabetize the list of bridesmaids and output the list and also the other women in the bridal party
Dim Pass As Integer
Dim Temp As String
Dim Slot As Integer
picWomen.Cls
picWomen.Print "Your Maid/Matron of Honor is "; MaidOfHonor; "."
picWomen.Print " "
picWomen.Print "Your Bridesmaids are: "
For Pass = 1 To NumberofBM - 1
    For Slot = 1 To NumberofBM - Pass
        If BridesmaidLast(Slot) > BridesmaidLast(Slot + 1) Then
            Temp = BridesmaidLast(Slot)
            BridesmaidLast(Slot) = BridesmaidLast(Slot + 1)
            BridesmaidLast(Slot + 1) = Temp
            Temp = BridesmaidFirst(Slot)
            BridesmaidFirst(Slot) = BridesmaidFirst(Slot + 1)
            BridesmaidFirst(Slot + 1) = Temp
        End If
    Next Slot
Next Pass
For Slot = 1 To NumberofBM
    picWomen.Print BridesmaidFirst(Slot); " "; BridesmaidLast(Slot)
Next Slot
picWomen.Print " "
picWomen.Print "Your Flower Girl is "; FlowerGirl; "."
picWomen.Print " "
picWomen.Print "Your Personal Attendent is "; PersonalAttendent; "."
End Sub

Private Sub cmdAlphaGMU_Click()
'this button is meant to alphabetize the groomsmen and alphabetize the ushers and display those lists and the other men in the wedding party.
picMen.Cls
Dim Slot As Integer
Dim Pass As Integer
Dim Temp As String
picMen.Print "Your Best Man is "; BestMan; "."
picMen.Print " "
picMen.Print "Your Groomsmen are: "
For Pass = 1 To NumberofGM - 1
    For Slot = 1 To NumberofGM - Pass
        If GroomsmenLast(Slot) > GroomsmenLast(Slot + 1) Then
            Temp = GroomsmenLast(Slot)
            GroomsmenLast(Slot) = GroomsmenLast(Slot + 1)
            GroomsmenLast(Slot + 1) = Temp
            Temp = GroomsmenFirst(Slot)
            GroomsmenFirst(Slot) = GroomsmenFirst(Slot + 1)
            GroomsmenFirst(Slot + 1) = Temp
        End If
    Next Slot
Next Pass
For Slot = 1 To NumberofGM
    picMen.Print GroomsmenFirst(Slot); " "; GroomsmenLast(Slot)
Next Slot
picMen.Print " "
picMen.Print "Your Ring Bearer is "; RingBearer; "."
picMen.Print " "
picMen.Print "Your Ushers are: "
For Pass = 1 To NumberofU - 1
    For Slot = 1 To NumberofU - Pass
        If UsherLast(Slot) > UsherLast(Slot + 1) Then
            Temp = BridesmaidLast(Slot)
            UsherLast(Slot) = UsherLast(Slot + 1)
            UsherLast(Slot + 1) = Temp
            Temp = UsherFirst(Slot)
            UsherFirst(Slot) = UsherFirst(Slot + 1)
            UsherFirst(Slot + 1) = Temp
        End If
    Next Slot
Next Pass
For Slot = 1 To NumberofU
    picMen.Print UsherFirst(Slot); " "; UsherLast(Slot)
Next Slot
End Sub

Private Sub cmdAttire_Click()
'this button is used to go to the form to select wedding party attire
frmAttire.Show
FrmWeddingParty.Hide
cmdshowAttireCost.Enabled = True
End Sub

Private Sub cmdBackTo_Click()
'this button is used to go back to the main menu
frmWedding.Show
FrmWeddingParty.Hide

End Sub

Private Sub cmdMen_Click()
Dim Slot As Integer
'this button is used to input the names of the men in your bridal party
BestMan = InputBox("Enter your Best Man's first and last name.", "Best Man's Name")
NumberofGM = InputBox("Enter the number of Groomsmen not including your Best Man.", "Number of Groomsmen")
For Slot = 1 To NumberofGM
    GroomsmenFirst(Slot) = InputBox("Enter the first name of one of your Groomsmen.", "Groomsmen First Name")
    GroomsmenLast(Slot) = InputBox("Enter the last name of the Groomsmen you just entered.", "Groomsmen Last Name")
Next Slot
RingBearer = InputBox("Enter your Ring Bearer's first and last name.", "Ring Bearer's Name")
NumberofU = InputBox("Enter the number of Ushers.", "Number of Ushers")
For Slot = 1 To NumberofU
    UsherFirst(Slot) = InputBox("Enter the first name of one of your Ushers.", "Ushers First Name")
    UsherLast(Slot) = InputBox("Enter the last name of the Usher you just entered.", "Ushers Last Name")
Next Slot
picMen.Print "Your Best Man is "; BestMan; "."
picMen.Print " "
picMen.Print "Your Groomsmen are: "
For Slot = 1 To NumberofGM
    picMen.Print GroomsmenFirst(Slot); " "; GroomsmenLast(Slot)
Next Slot
picMen.Print " "
picMen.Print "Your Ring Bearer is "; RingBearer; "."
picMen.Print " "
picMen.Print "Your Ushers are: "
For Slot = 1 To NumberofU
    picMen.Print UsherFirst(Slot); " "; UsherLast(Slot)
Next Slot
cmdAttire.Enabled = True
cmdAlphaGMU.Enabled = True
cmdMen.Enabled = False
End Sub

Private Sub cmdQuit_Click()
'this button is to end the program
End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub cmdshowAttireCost_Click()
'this button is used to display the total cost of the attire
If TotalCostofAttire = 0 Then
    TotalCostofAttire = CostofDress + TotalCostBM + CostofTux + TotalCostGMU
End If
picAttire.Print "The Total Cost of the Wedding Party Attire is "; FormatCurrency(TotalCostofAttire); "."

End Sub

Private Sub cmdWomen_Click()
'this button is used to input the names of the women in the bridal party
Dim Slot As Integer
MaidOfHonor = InputBox("Enter your Matron/Maid of Honor's first and last name.", "Maid/Matron of Honor's Name")
NumberofBM = InputBox("Enter the number of Bridesmaids not including your Matron/Maid of Honor.", "Number of Bridesmaids")
For Slot = 1 To NumberofBM
   BridesmaidFirst(Slot) = InputBox("Enter the first name of one of your Bridesmaids.", "Bridesmaids First Name")
   BridesmaidLast(Slot) = InputBox("Enter the last name of the bridesmaid that you just entered.", "Bridesmaids Last Name")
Next Slot
FlowerGirl = InputBox("Enter your Flower Girl's first and last name.", "Flower Girl's Name")
PersonalAttendent = InputBox("Enter your Personal Attendent's first and last name.", "Personal Attendent's Name")
picWomen.Print "Your Maid/Matron of Honor is "; MaidOfHonor; "."
picWomen.Print " "
picWomen.Print "Your Bridesmaids are: "
For Slot = 1 To NumberofBM
    picWomen.Print BridesmaidFirst(Slot); " "; BridesmaidLast(Slot)
Next Slot
picWomen.Print " "
picWomen.Print "Your Flower Girl is "; FlowerGirl; "."
picWomen.Print " "
picWomen.Print "Your Personal Attendent is "; PersonalAttendent; "."
cmdMen.Enabled = True
cmdAlphaBM.Enabled = True
cmdWomen.Enabled = False
End Sub

