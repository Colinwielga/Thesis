VERSION 5.00
Begin VB.Form frmHunting 
   BackColor       =   &H000080FF&
   Caption         =   "Load Hunting Prices and Go to Hunting Page"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   ScaleHeight     =   7890
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNonResidentFind 
      BackColor       =   &H0000FF00&
      Caption         =   "Non-Resident Click Here"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdNonResidentLoad 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here To Load Non-Resident Licensing Information"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdDuckandPhesant 
      BackColor       =   &H0000FF00&
      Caption         =   "Duck and Pheasant Hunters Click here for Important Information"
      Height          =   1455
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   5415
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   3600
      Picture         =   "frmHunting.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   3480
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   6720
      Picture         =   "frmHunting.frx":0DB5
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   480
      Picture         =   "frmHunting.frx":1D26
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdRandom 
      BackColor       =   &H0000FF00&
      Caption         =   "Click here for Random Facts About Hunting In Minnesota"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   5415
   End
   Begin VB.CommandButton cmdSortPrice 
      BackColor       =   &H0000FF00&
      Caption         =   "Sort the Licenses by Price"
      Enabled         =   0   'False
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000C&
      Height          =   2895
      Left            =   3360
      ScaleHeight     =   2835
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H0000FF00&
      Caption         =   "Display the Animal and the Price"
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmHunting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Outdoors
'Hunting
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this is the hunting page, it tells how much the licenses costs and has random facts
'declaring global variables
Dim AnimalNonResident(1 To 50) As String, NonResidentPrice(1 To 50) As Single
Dim NonCtr As Integer


Private Sub cmdDuckandPhesant_Click()

MsgBox "Duck and pheasant hunting require you to buy a stamp for $7.50 as-well-as a small game license", , "Attention"

End Sub



Private Sub cmdNonResidentFind_Click()

'declaring variables
Dim j As Integer, Ctr As Integer, found As Boolean
Dim x As String

picResults.Cls

x = InputBox("Enter the License you want to purchase (Enter Animal with Capital First Letter example Bear)")

For j = 1 To NonCtr
    If AnimalNonResident(j) = x Then
        picResults.Print "The Non-Resident Price for "; AnimalNonResident(j); " is "; FormatCurrency(NonResidentPrice(j), 2)
        found = True
    End If
Next j

If Not found Then
    picResults.Print "That animal is not hunted in Minnesota or allowed to be "
    picResults.Print "Hunted by Non-Residents"
    picResults.ForeColor = vbGreen
End If
    
End Sub

Private Sub cmdNonResidentLoad_Click()

Open App.Path & "\NonResident.txt" For Input As #1

NonCtr = 0

Do While Not EOF(1)
    NonCtr = NonCtr + 1
    Input #1, AnimalNonResident(NonCtr), NonResidentPrice(NonCtr)
Loop

Close #1

cmdNonResidentLoad.Enabled = False
cmdNonResidentFind.Enabled = True

End Sub

Private Sub cmdRandom_Click()
'declaring variables
Dim Facts(1 To 50) As String, Ctr As Integer

Ctr = 0

Open App.Path & "\HuntingRandomFacts.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Facts(Ctr)
Loop

Close #1
picResults.Cls
picResults.Print Facts(CInt(Int((6 * Rnd()) + 1)))
picResults.ForeColor = vbGreen

End Sub

Private Sub cmdReturn_Click()

frmDNR.Show
frmHunting.Hide

End Sub

Private Sub cmdSortPrice_Click()

'clearing the picture box
picResults.Cls


picResults.Print "Animal"; Tab(30); "Licenses Price"
picResults.Print "------------------------------------------------------------------------"
picResults.ForeColor = vbGreen

'declaring variables
Dim pos As Integer, pass As Integer, tempAnimal As String, tempPrice As Single
Dim j As Integer

'Sorting the array by price
For pass = 1 To HuntingCtr - 1
    For pos = 1 To HuntingCtr - pass
        If HuntingPrice(pos) < HuntingPrice(pos + 1) Then
            tempPrice = HuntingPrice(pos)
            HuntingPrice(pos) = HuntingPrice(pos + 1)
            HuntingPrice(pos + 1) = tempPrice
            tempAnimal = Animal(pos)
            Animal(pos) = Animal(pos + 1)
            Animal(pos + 1) = tempAnimal
        End If
    Next pos
Next pass

For j = 1 To HuntingCtr
    picResults.Print Animal(j); Tab(30); FormatCurrency(HuntingPrice(j), 2)
    picResults.ForeColor = vbGreen
Next j

End Sub

Private Sub cmdDisplay_Click()
'clearing the picture box
picResults.Cls

'declaring variables
Dim I As Integer

picResults.ForeColor = vbGreen
picResults.Print "Animal"; Tab(30); "Licenses Price"
picResults.Print "------------------------------------------------------------------------"


'reading the prices into the picture box
For I = 1 To HuntingCtr
    picResults.Print Animal(I); Tab(30); FormatCurrency(HuntingPrice(I), 2)
    picResults.ForeColor = vbGreen
Next I

cmdSortPrice.Enabled = True
cmdDisplay.Enabled = False

End Sub

Private Sub Image1_Click()

End Sub
