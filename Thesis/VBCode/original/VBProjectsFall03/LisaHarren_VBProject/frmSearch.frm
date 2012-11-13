VERSION 5.00
Begin VB.Form InventorySearch 
   BackColor       =   &H000040C0&
   Caption         =   "Inventory Search"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit2 
      Caption         =   "Back to Main Menu"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox pbxResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   2280
      ScaleHeight     =   2835
      ScaleWidth      =   7755
      TabIndex        =   2
      Top             =   480
      Width           =   7815
   End
   Begin VB.CommandButton cmdSearchCost 
      Caption         =   "Seach by Purchasing Cost"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearchName 
      Caption         =   "Search by Name"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "InventorySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SalesReport (VBProject.vbp)
'InventorySeach (frmSearch.frm)
'Written by Lisa Harren
'10-23-03 (finished)
'Form Purpose:  The purpose of this form is to allow the user to search for an
      'inventory item, and its information, by either name or puchasing price.
      

Option Explicit
'Forces the writer to declare any variables before they may be used.


Private Sub cmdClear2_Click()
'Clears anything that may be in pxbResults2

pbxResults2.Cls
End Sub


Private Sub cmdQuit2_Click()
'Allows user to move from the Search Form back to the Main Menu.

  InventorySearch.Hide
  MainMenu.Show
  
End Sub

Private Sub cmdSearchCost_Click()
Dim Found As Boolean
Dim NewNumber As Single
'Seaches the Inventory for an Item by using the Item's purchasing price.
NewNumber = InputBox("Enter Item Puchase Price", "Price Search")
Found = False
i = 0
Do Until Found Or i = 5
  i = i + 1
  If NewNumber = pPrice(i) Then
    Found = True
  End If
  pbxResults2.Cls
Loop

'Prints the results.
If Found = True Then
  pbxResults2.Print "Item", "Items Sold", "Items Remaining", "Purchase Price", "Retail Price"
  pbxResults2.Print strName(i), sold(i), number(i), , FormatCurrency(pPrice(i), 2), , FormatCurrency(sPrice(i), 2)
 Else
  pbxResults2.Print "Item Not Found!  Please Check for Input Errors."
End If

End Sub

Private Sub cmdSearchName_Click()
Dim Found As Boolean
Dim NewName As String

'Searches the Inventory by using the item's name.
NewName = InputBox("Enter Item Name(Capitalizing the First Letter)", "Item Search")
Found = False
i = 0
Do Until Found Or i = 5
  i = i + 1
  If NewName = strName(i) Then
    Found = True
  End If
pbxResults2.Cls
Loop

'Prints the results.
If Found = True Then
  pbxResults2.Print "Item", "Items Sold", "Items Remaining", "Purchase Price", "Retail Price"
  pbxResults2.Print strName(i), sold(i), number(i), , FormatCurrency(pPrice(i), 2), , FormatCurrency(sPrice(i), 2)
 Else
  pbxResults2.Print "Item Not Found! Please Check For Spelling Errors!"
End If

End Sub

Private Sub Form_Load()
'Opens the file for use in the program.

strPath = "N:\CS130\handin\"
Dim strFile As String
strFile = strPath & "LisaHarren_VBProject\Items.txt"
Open strFile For Input As #1
For i = 1 To 5
  Input #1, strName(i), sold(i), number(i), pPrice(i), sPrice(i)
Next i
Close #1
  

End Sub
