VERSION 5.00
Begin VB.Form MainMenu 
   BackColor       =   &H80000001&
   Caption         =   "Main Menu"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxResults 
      BackColor       =   &H8000000E&
      Height          =   3375
      Left            =   2280
      ScaleHeight     =   3315
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Finish"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Seach Inventory"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdInventoryDisplay 
      Caption         =   "Display Inventory"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrintSales 
      Caption         =   "Print Sales Stats"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdRetrieve 
      BackColor       =   &H8000000D&
      Caption         =   "Get Day's Numbers"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SalesReport (VBProject.vbp)
'Main Menu (frmMain)
'Written by Lisa Harren
'10-23-03 (Finished)
'Project Purpose:  The Purpose of this project is to aide a Department
      'Store in keeping track of their Inventory and Sales each day.
'Form Purpose:  The purpose of this form is to allow the user to access
      'the sales, revenues, gains, losses, and Inventory cound for the day.
      


Option Explicit
'Forces you to declare all variables before they may be used.



Private Sub cmdClear_Click()
'Clears anything that may be in pbxResults.

pbxResults.Cls

End Sub



Private Sub cmdInventoryDisplay_Click()

pbxResults.Cls
'Automatically Clears anything that may be in pbxResults

Dim ctemp As Integer, stemp As Integer
Dim ptemp As Single, xtemp As Single
Dim temp As String
Dim Pass As Integer, N As Integer

'Display Inventory by how many items remain at the end of the day.
N = 5
For Pass = 1 To N
  For i = 1 To (Pass - 1)
    If number(i) < number(i + 1) Then
      ctemp = number(i)
      number(i) = number(i + 1)
      number(i + 1) = ctemp
      temp = strName(i)
      strName(i) = strName(i + 1)
      strName(i + 1) = temp
      stemp = sold(i)
      sold(i) = sold(i + 1)
      sold(i + 1) = stemp
      ptemp = pPrice(i)
      pPrice(i) = pPrice(i + 1)
      pPrice(i + 1) = ptemp
      xtemp = sPrice(i)
      sPrice(i) = sPrice(i + 1)
      sPrice(i + 1) = xtemp
    End If
  Next i
Next Pass

'Print the above sort.
pbxResults.Print "Item", "# Items Remaining"
For i = 1 To 5
  pbxResults.Print strName(i), number(i)
Next i

'Display a message box to remind of re-order needed for any inventory of less than 20 items.
MsgBox "Please re-order any inventory item with less than 20 items in stock", , "Reminder"


End Sub



Private Sub cmdPrintSales_Click()

pbxResults.Cls
'Automatically clears anything that may be in pbxResults.

Dim R(1 To 5) As Double, S As Single, C As Single, P As Single

'Calcultates and Prints the Revenue for the day for each inventory item.
pbxResults.Print "Item", "# Items Sold", "Item Revenue"
For i = 1 To 5
  R(i) = sold(i) * sPrice(i)
  pbxResults.Print strName(i), sold(i), FormatCurrency(R(i), 2)
Next i

pbxResults.Print
'Places an empty line in pbxResults.

'Prints the Totoal Sales Revenue for the day.
For i = 1 To 5
  S = S + R(i)
Next i
pbxResults.Print "Total Sales Revenue for today was "; FormatCurrency(S, 2)

pbxResults.Print
'Places an empty line in pxbResults.

'Prints the Net Sales for the day.
For i = 1 To 5
  C = C + pPrice(i) * sold(i)
Next i
P = S - C
pbxResults.Print "The Net Sales for today were "; FormatCurrency(P, 2)

End Sub

Private Sub cmdQuit_Click()
'Takes user from the Main Menu to the Good Night form.
GoodNight.Show
MainMenu.Hide

End Sub
'Opens the file link.
Private Sub cmdRetrieve_Click()
strPath = "N:\CS130\handin\"
Dim strFile As String
strFile = strPath & "LisaHarren_VBProject\Items.txt"
Open strFile For Input As #1
For i = 1 To 5
  Input #1, strName(i), sold(i), number(i), pPrice(i), sPrice(i)
Next i
Close #1

End Sub

Private Sub cmdSearch_Click()
'Takes the user from the Main Menu to the Search form.
MainMenu.Hide
InventorySearch.Show

End Sub

